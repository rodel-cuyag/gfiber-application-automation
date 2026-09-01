"""
eod_report.py
--------------
Builds the "EOD Report" sheet: aggregate KPIs for a calling-day RANGE
(start_date..end_date, inclusive — a single day is just a range of 1),
covering the full funnel the GFiber abandoned-application agent measures:
identity -> consent -> postpaid -> intent -> endorsement.

Only metrics that can be honestly derived from the available data are
computed. Metrics the source data can't support (e.g. LLM inference
cost) are left blank with a comment, rather than guessed.
"""

import pandas as pd


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

    def count_by_postpaid(column, value, postpaid_value):
        return (
            (connected_calls[column] == value)
            & (connected_calls["Postpaid Status"] == postpaid_value)
        ).sum()

    contacted = count("Participated Call", "Yes")

    identity_confirmed = count("Identity Confirmed", "Yes")
    wrong_customer = count("Identity Confirmed", "No")
    consented = count("Consent & Recording Confirmed", "Yes")
    declined_recording = count("Consent & Recording Confirmed", "No")
    postpaid = count("Postpaid Status", "postpaid")
    non_postpaid = count("Postpaid Status", "non_postpaid")

    wishes_to_proceed = count("Application Intent", "proceed")
    no_longer_interested = count("Application Intent", "no_longer_interested")
    already_completed = count("Application Intent", "already_completed")

    # Intent broken down by customer segment — an Application Intent x
    # Postpaid Status cross-tab, deliberately independent of the
    # final_disposition-derived Postpaid/Non-Postpaid Conversion rows below:
    # the two pairs answer different questions and may legitimately differ.
    # These sub-counts won't always sum back to their total, since a call
    # with a blank Postpaid Status falls into neither segment.
    wishes_to_proceed_postpaid = count_by_postpaid("Application Intent", "proceed", "postpaid")
    wishes_to_proceed_non_postpaid = count_by_postpaid("Application Intent", "proceed", "non_postpaid")
    no_longer_interested_postpaid = count_by_postpaid("Application Intent", "no_longer_interested", "postpaid")
    no_longer_interested_non_postpaid = count_by_postpaid("Application Intent", "no_longer_interested", "non_postpaid")
    already_completed_postpaid = count_by_postpaid("Application Intent", "already_completed", "postpaid")
    already_completed_non_postpaid = count_by_postpaid("Application Intent", "already_completed", "non_postpaid")

    endorsed = count("Endorsed for Work Order", "Yes")
    lead_outbound = count("Lead for Outbound Handling", "Yes")
    lead_email = count("Lead for Email Remarketing", "Yes")

    postpaid_conversion = count("Final Disposition", "postpaid_wishes_to_proceed")
    non_postpaid_conversion = count("Final Disposition", "non_postpaid_wishes_to_proceed")
    not_available_no_consent = count("Final Disposition", "not_available_no_consent")

    non_completion_price = count("Non-Completion Reason", "price")
    non_completion_competitor = count("Non-Completion Reason", "competitor")
    competitor_identified = count("Competitor Detected", "Yes")

    repeat_requested = count("Repeat Requested", "Yes")
    identity_reasked = count("Identity Re-asked (defect)", "Yes")
    opt_out = count("Opt-Out Flag", "Yes")

    connection_rate = round((connected / dialed) * 100, 1) if dialed else 0.0
    conversion_rate = round((wishes_to_proceed / connected) * 100, 1) if connected else 0.0

    # Share-of-connected percentages. Connected is the right denominator for
    # all of these: every KPI-derived count above is already connected-only,
    # so the base and the numerator come from the same population.
    def pct_of_connected(n):
        return round((n / connected) * 100, 1) if connected else 0.0

    postpaid_pct = pct_of_connected(postpaid)
    non_postpaid_pct = pct_of_connected(non_postpaid)
    endorsed_pct = pct_of_connected(endorsed)
    lead_outbound_pct = pct_of_connected(lead_outbound)
    lead_email_pct = pct_of_connected(lead_email)

    # Calculate durations
    durations = range_log["Call Duration (sec)"].dropna()
    avg_duration = round(durations.mean(), 1) if not durations.empty else None

    total_duration_sec = durations.sum() if not durations.empty else 0
    total_duration_min = round(total_duration_sec / 60, 1) if total_duration_sec else None

    # Calculate retries queued (Failed, No Answer, Busy = need retry)
    retries_queued = failed + no_answer + busy

    metrics = [
        ("Report Period", period_label),
        ("Days in Range", days_in_range),
        ("Agent ID", agent_id),
        ("", ""),  # Blank row for readability

        # Call Volume Metrics
        ("Calls Dialed - Target", ""),
        ("Calls Dialed - Actual", dialed),
        ("Calls Connected", connected),
        ("No Answer", no_answer),
        ("Busy", busy),
        ("Failed", failed),
        ("", ""),  # Blank row

        # Participation
        ("Contacted", contacted),
        ("Total Completed Calls", completed),
        ("", ""),  # Blank row

        # Duration Metrics
        ("Total Call Duration (minutes)", total_duration_min),
        ("Avg. Call Duration - Connected (seconds)", avg_duration),
        ("", ""),  # Blank row

        # Funnel
        ("Identity Confirmed", identity_confirmed),
        ("Wrong Customer", wrong_customer),
        ("Consented to Continue and Recording", consented),
        ("Declined Recording", declined_recording),
        ("Postpaid Verified", postpaid),
        ("Postpaid Customers % (of Connected)", f"{postpaid_pct}%"),
        ("Non-Postpaid Verified", non_postpaid),
        ("Non-Postpaid Customers % (of Connected)", f"{non_postpaid_pct}%"),
        ("", ""),  # Blank row

        # Intent Outcomes
        ("Wishes to Proceed", wishes_to_proceed),
        ("Wishes to Proceed - Postpaid", wishes_to_proceed_postpaid),
        ("Wishes to Proceed - Non-Postpaid", wishes_to_proceed_non_postpaid),
        ("No Longer Interested", no_longer_interested),
        ("No Longer Interested - Postpaid", no_longer_interested_postpaid),
        ("No Longer Interested - Non-Postpaid", no_longer_interested_non_postpaid),
        ("Application Already Completed", already_completed),
        ("Application Already Completed - Postpaid", already_completed_postpaid),
        ("Application Already Completed - Non-Postpaid", already_completed_non_postpaid),
        ("", ""),  # Blank row

        # Endorsement & Leads
        ("Endorsed for Work Order", endorsed),
        ("Endorsed for Work Order % (of Connected)", f"{endorsed_pct}%"),
        ("Lead for Outbound Handling", lead_outbound),
        ("Lead for Outbound Handling % (of Connected)", f"{lead_outbound_pct}%"),
        ("Email Remarketing Tagged", lead_email),
        ("Email Remarketing % (of Connected)", f"{lead_email_pct}%"),
        ("", ""),  # Blank row

        # Final Dispositions
        ("Postpaid Conversion", postpaid_conversion),
        ("Non-Postpaid Conversion", non_postpaid_conversion),
        ("Not Available / No Consent", not_available_no_consent),
        ("", ""),  # Blank row

        # Conversion Metrics
        ("Connection Rate (Connected / Dialed)", f"{connection_rate}%"),
        ("Conversion Rate (Proceed / Connected)", f"{conversion_rate}%"),
        ("Retries Queued for Tomorrow", retries_queued),
        ("", ""),  # Blank row

        # Non-Completion & Competitor
        ("Non-Completion - Price", non_completion_price),
        ("Non-Completion - Competitor", non_completion_competitor),
        ("Competitor Identified", competitor_identified),
        ("", ""),  # Blank row

        # Quality
        ("Repeat Requested (quality)", repeat_requested),
        ("Identity Re-asked (defect)", identity_reasked),
        ("Opt-Out Requested", opt_out),
        ("", ""),  # Blank row

        # FINOPS Section
        ("FINOPS", ""),
        ("LLM Inference Cost (USD)", ""),
        ("Total Daily Spend (USD)", ""),
        ("", ""),  # Blank row

        # ISSUES & CHANGES Section
        ("ISSUES & CHANGES", ""),
        ("Open P0 Issues", ""),
        ("Open P1 Issues", ""),
        ("Changes Deployed Today", ""),
        ("Changes Pending Approval for Tomorrow", ""),
        ("", ""),  # Blank row

        # TOMORROW'S PLAN Section
        ("TOMORROW'S PLAN", ""),
        ("Target Call Volume", ""),
        ("Expected List from Globe (ETA)", ""),
        ("Calling Window", "9:00 AM - 6:00 PM PHT"),
        ("Phase Gate Status", ""),
    ]

    return pd.DataFrame(metrics, columns=["Metric", "Value"])
