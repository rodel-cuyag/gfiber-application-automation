"""
Builds the "EOD Report" sheet: aggregate KPIs for a calling-day RANGE
(start_date..end_date, inclusive — a single day is just a range of 1),
covering the full funnel the GFiber abandoned-application agent measures:
identity -> consent -> postpaid -> intent -> endorsement.

Only metrics that can be honestly derived from the available data are
computed. Metrics the source data can't support (e.g. LLM inference
cost) are left blank with a comment, rather than guessed.
"""

import pandas as pd

# Competitor breakdown metric keys are built at runtime (one per provider
# actually seen), so they can't live in a static list. excel_writer finds
# them by this prefix; the emission order below is the sheet order.
COMPETITOR_METRIC_PREFIX = "Competitor - "

# Display order and labels for the providers the agent config enumerates
# for competitor_name. 'none' is excluded (it means the customer did not
# switch at all) and 'others' is emitted near the end as the reconciling
# residual rather than listed here.
CANONICAL_COMPETITORS = [
    ("pldt", "PLDT"),
    ("converge", "Converge"),
    ("starlink", "Starlink"),
    ("sky", "Sky"),
    ("dito", "DITO"),
    ("smart", "Smart"),
]

# Closes the breakdown, below even the "Others" residual, rather than
# sitting among the canonical providers above: 'undisclosed' names no
# provider, so it doesn't belong in the list of real competitors. Still a
# named row of its own (unlike the values in _NON_PROVIDER_VALUES) — the
# count is meaningful, and "Others" is computed net of it.
TRAILING_COMPETITORS = [
    ("undisclosed", "Undisclosed"),
]

# Values that never earn a named row of their own: 'none' means the field
# didn't apply, and 'others' is the residual bucket itself.
_NON_PROVIDER_VALUES = {"none", "others", ""}


def _build_competitor_rows(competitor_calls) -> list:
    """
    One (metric key, count) pair per competitor, in display order:
    the canonical providers first (always present, zeros included, so the
    row set stays stable day to day for the Yesterday/Δ lookup), then a
    row for any write-in provider found in the data, then "Others", then
    "Undisclosed".

    The agent config tells the model to return an unlisted provider's name
    verbatim rather than "others", so the value set is open-ended — the
    write-in rows are what stop those calls from vanishing. "Others" is
    computed as the residual so the breakdown always reconciles to
    Competitor Identified, absorbing the literal 'others' value along with
    any detected call whose competitor_name came back blank or 'none'.
    """
    names = (
        competitor_calls["Competitor Name"]
        .dropna()
        .astype(str)
        .str.strip()
        .str.lower()
    )
    counts = names.value_counts()

    def named_rows(spec):
        return [(label, int(counts.get(key, 0))) for key, label in spec]

    known = (
        {key for key, _ in CANONICAL_COMPETITORS}
        | {key for key, _ in TRAILING_COMPETITORS}
        | _NON_PROVIDER_VALUES
    )
    extras = [(k, int(v)) for k, v in counts.items() if k not in known]
    extras.sort(key=lambda kv: (-kv[1], kv[0]))

    leading = named_rows(CANONICAL_COMPETITORS) + [(k.title(), v) for k, v in extras]
    trailing = named_rows(TRAILING_COMPETITORS)

    # Net of the trailing rows as well as the leading ones, even though it
    # prints above them — otherwise Others would double-count those calls
    # and the breakdown would overshoot Competitor Identified.
    claimed = sum(v for _, v in leading) + sum(v for _, v in trailing)
    others = ("Others", len(competitor_calls) - claimed)

    rows = leading + [others] + trailing

    return [(f"{COMPETITOR_METRIC_PREFIX}{label}", value) for label, value in rows]


def build_eod_report(call_detail_log: pd.DataFrame, start_date, end_date, agent_id: int) -> pd.DataFrame:
    """
    Filters the call detail log to [start_date, end_date] (inclusive) and
    returns a single aggregated key/value summary DataFrame covering the
    whole period.
    """
    range_log = call_detail_log[
        (call_detail_log["Call Date (PHT)"] >= start_date)
        & (call_detail_log["Call Date (PHT)"] <= end_date)
    ]

    days_in_range = (end_date - start_date).days + 1
    period_label = str(start_date) if start_date == end_date else f"{start_date} to {end_date}"

    dialed = len(range_log)

    # Status comes exclusively from the Twilio call-progress journey
    # (see call_detail.py) and is mapped to display-friendly values:
    # "Connected", "Failed", "No Answer", "Busy", etc.
    connected = (range_log["Status"] == "Connected").sum()
    failed = (range_log["Status"] == "Failed").sum()
    no_answer = (range_log["Status"] == "No Answer").sum()
    busy = (range_log["Status"] == "Busy").sum()

    # Completed comes from the KPI-derived call_completed flag (see
    # call_detail.py's "Call Completed" column) — distinct from "Connected",
    # which is the Twilio call-progress outcome.
    completed = (range_log["Call Completed"] == "Yes").sum()

    # Every KPI-derived count is taken from connected calls only, so the
    # funnel and the rates built on it stay honest — a no-answer call can't
    # have confirmed an identity or stated an intent.
    connected_calls = range_log[range_log["Status"] == "Connected"]

    def count(column, value):
        return (connected_calls[column] == value).sum()

    # The funnel is strictly nested: each level is a subset of the one above,
    # so the sheet's "X Breakdown" banners describe a real decomposition
    # rather than independent counts that happen to sit near each other.
    #
    # Note the top level doesn't fully reconcile: Identity Confirmed +
    # Wrong Customer is less than Calls Connected, because a connected call
    # with no meaningful engagement (or no KPI record at all) never reaches
    # the identity step. That residual is deliberately not shown as its own
    # row — see the Validation Report's funnel-residual audit step.
    identity_calls = connected_calls[connected_calls["Identity Confirmed"] == "Yes"]
    identity_confirmed = len(identity_calls)
    wrong_customer = (connected_calls["Identity Confirmed"] == "No").sum()

    consented_calls = identity_calls[identity_calls["Consent & Recording Confirmed"] == "Yes"]
    consented = len(consented_calls)
    declined_recording = (identity_calls["Consent & Recording Confirmed"] == "No").sum()

    postpaid_calls = consented_calls[consented_calls["Postpaid Status"] == "postpaid"]
    non_postpaid_calls = consented_calls[consented_calls["Postpaid Status"] == "non_postpaid"]
    postpaid = len(postpaid_calls)
    non_postpaid = len(non_postpaid_calls)

    wishes_to_proceed = count("Application Intent", "proceed")
    no_longer_interested = count("Application Intent", "no_longer_interested")
    already_completed = count("Application Intent", "already_completed")

    # Intent within each customer segment. Deliberately independent of the
    # final_disposition-derived Postpaid/Non-Postpaid Conversion rows below:
    # the two pairs answer different questions and may legitimately differ.
    def intent(segment_calls, value):
        return (segment_calls["Application Intent"] == value).sum()

    wishes_to_proceed_postpaid = intent(postpaid_calls, "proceed")
    wishes_to_proceed_non_postpaid = intent(non_postpaid_calls, "proceed")
    no_longer_interested_postpaid = intent(postpaid_calls, "no_longer_interested")
    no_longer_interested_non_postpaid = intent(non_postpaid_calls, "no_longer_interested")
    already_completed_postpaid = intent(postpaid_calls, "already_completed")
    already_completed_non_postpaid = intent(non_postpaid_calls, "already_completed")

    endorsed = count("Endorsed for Work Order", "Yes")
    lead_outbound = count("Lead for Outbound Handling", "Yes")
    lead_email = count("Lead for Email Remarketing", "Yes")

    postpaid_conversion = count("Final Disposition", "postpaid_wishes_to_proceed")
    non_postpaid_conversion = count("Final Disposition", "non_postpaid_wishes_to_proceed")
    not_available_no_consent = count("Final Disposition", "not_available_no_consent")

    non_completion_price = count("Non-Completion Reason", "price")
    non_completion_competitor = count("Non-Completion Reason", "competitor")
    competitor_calls = connected_calls[connected_calls["Competitor Detected"] == "Yes"]
    competitor_identified = len(competitor_calls)
    competitor_rows = _build_competitor_rows(competitor_calls)

    repeat_requested = count("Repeat Requested", "Yes")
    identity_reasked = count("Identity Re-asked (defect)", "Yes")
    opt_out = count("Opt-Out Flag", "Yes")

    connection_rate = round((connected / dialed) * 100, 1) if dialed else 0.0
    conversion_rate = round((wishes_to_proceed / connected) * 100, 1) if connected else 0.0

    def pct_of(n, denominator):
        return round((n / denominator) * 100, 1) if denominator else 0.0

    # Share-of-connected percentages for the outcome metrics: every
    # KPI-derived count is already connected-only, so the base and the
    # numerator come from the same population.
    def pct_of_connected(n):
        return pct_of(n, connected)

    # The funnel pair is a share of its immediate parent instead, so the
    # two sum to 100% within the "Consented" breakdown they sit under.
    postpaid_pct = pct_of(postpaid, consented)
    non_postpaid_pct = pct_of(non_postpaid, consented)
    endorsed_pct = pct_of_connected(endorsed)
    lead_outbound_pct = pct_of_connected(lead_outbound)
    lead_email_pct = pct_of_connected(lead_email)

    # Over Connected, the same base as the headline conversion rate, so the
    # two decompose it. That decomposition holds only while final_disposition
    # and application_intent agree — they're separate agent fields, so if the
    # two segment rates stop summing to the headline, that's a data-quality
    # signal about the agent's output, not a fault in the sheet.
    postpaid_conversion_rate = pct_of_connected(postpaid_conversion)
    non_postpaid_conversion_rate = pct_of_connected(non_postpaid_conversion)

    # Calculate durations
    durations = range_log["Call Duration (sec)"].dropna()
    avg_duration = round(durations.mean(), 1) if not durations.empty else None

    total_duration_sec = durations.sum() if not durations.empty else 0
    total_duration_min = round(total_duration_sec / 60, 1) if total_duration_sec else None

    # Calculate retries queued (Failed, No Answer, Busy)
    retries_queued = failed + no_answer + busy

    metrics = [
        ("Report Period", period_label),
        ("Days in Range", days_in_range),
        ("Agent ID", agent_id),
        ("", ""),

        # Call Volume Metrics
        ("Calls Dialed - Target", ""),
        ("Calls Dialed - Actual", dialed),
        ("Calls Connected", connected),
        ("No Answer", no_answer),
        ("Busy", busy),
        ("Failed", failed),
        ("", ""),

        # Participation
        ("Total Completed Calls", completed),
        ("", ""),

        # Duration Metrics
        ("Total Call Duration (minutes)", total_duration_min),
        ("Avg. Call Duration - Connected (seconds)", avg_duration),
        ("", ""),

        # Funnel
        ("Identity Confirmed", identity_confirmed),
        ("Wrong Customer", wrong_customer),
        ("Consented to Continue and Recording", consented),
        ("Declined Recording", declined_recording),
        ("Postpaid Verified", postpaid),
        ("Postpaid Customers % (of Consented)", f"{postpaid_pct}%"),
        ("Non-Postpaid Verified", non_postpaid),
        ("Non-Postpaid Customers % (of Consented)", f"{non_postpaid_pct}%"),
        ("", ""),

        # Intent within each customer segment (shown under the FUNNEL
        # section's Postpaid / Non-Postpaid breakdown banners)
        ("Wishes to Proceed - Postpaid", wishes_to_proceed_postpaid),
        ("No Longer Interested - Postpaid", no_longer_interested_postpaid),
        ("Application Already Completed - Postpaid", already_completed_postpaid),
        ("Wishes to Proceed - Non-Postpaid", wishes_to_proceed_non_postpaid),
        ("No Longer Interested - Non-Postpaid", no_longer_interested_non_postpaid),
        ("Application Already Completed - Non-Postpaid", already_completed_non_postpaid),
        ("", ""),

        # Intent Outcomes (combined)
        ("Wishes to Proceed", wishes_to_proceed),
        ("No Longer Interested", no_longer_interested),
        ("Application Already Completed", already_completed),
        ("", ""),

        # Endorsement & Leads
        ("Endorsed for Work Order", endorsed),
        ("Endorsed for Work Order % (of Connected)", f"{endorsed_pct}%"),
        ("Lead for Outbound Handling", lead_outbound),
        ("Lead for Outbound Handling % (of Connected)", f"{lead_outbound_pct}%"),
        ("Email Remarketing Tagged", lead_email),
        ("Email Remarketing % (of Connected)", f"{lead_email_pct}%"),
        ("", ""),

        # Final Dispositions. No longer rendered on the EOD Report sheet
        # (see excel_writer.DASHBOARD_ROWS) — retained here because the
        # Validation Report's Calculation Audit still checks them against
        # its own independent recomputation.
        ("Not Available / No Consent", not_available_no_consent),
        ("", ""),

        # Conversion Metrics
        ("Connection Rate (Connected / Dialed)", f"{connection_rate}%"),
        ("Conversion Rate (Proceed / Connected)", f"{conversion_rate}%"),
        ("Postpaid Conversion Rate", f"{postpaid_conversion_rate}%"),
        ("Non-Postpaid Conversion Rate", f"{non_postpaid_conversion_rate}%"),
        # Validation-Report-only, same as Not Available / No Consent above.
        ("Retries Queued for Tomorrow", retries_queued),
        ("", ""),

        # Non-Completion & Competitor
        ("Non-Completion - Price", non_completion_price),
        ("Non-Completion - Competitor", non_completion_competitor),
        ("Competitor Identified", competitor_identified),
        ("", ""),

        # Competitor breakdown — variable length, one row per provider
        *competitor_rows,
        ("", ""),

        # Quality
        ("Repeat Requested (quality)", repeat_requested),
        ("Identity Re-asked (defect)", identity_reasked),
        ("Opt-Out Requested", opt_out),
        ("", ""),

        # FINOPS Section
        ("FINOPS", ""),
        ("LLM Inference Cost (USD)", ""),
        ("Total Daily Spend (USD)", ""),
        ("", ""),

        # ISSUES & CHANGES Section
        ("ISSUES & CHANGES", ""),
        ("Open P0 Issues", ""),
        ("Open P1 Issues", ""),
        ("Changes Deployed Today", ""),
        ("Changes Pending Approval for Tomorrow", ""),
        ("", ""),

        # TOMORROW'S PLAN Section
        ("TOMORROW'S PLAN", ""),
        ("Target Call Volume", ""),
        ("Expected List from Globe (ETA)", ""),
        ("Calling Window", "9:00 AM - 6:00 PM PHT"),
        ("Phase Gate Status", ""),
    ]

    return pd.DataFrame(metrics, columns=["Metric", "Value"])
