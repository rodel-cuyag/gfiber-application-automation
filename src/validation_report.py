"""
validation_report.py
--------------------
Builds a multi-sheet validation workbook that audits the EOD report
pipeline: source-data join coverage, per-row field completeness, a
step-by-step calculation trace, and a consolidated data-quality issues
register.

Generated automatically alongside every EOD report run.
"""

import json

import pandas as pd
from src import call_detail, data_loader


# ── Sheet 1: Join Summary ─────────────────────────────────────────

def _build_join_summary(working_table, agent_id, start_date, end_date):
    raw = data_loader.load_all()

    conv = raw["conversations"]
    conv_agt = conv[conv["agent_id"] == agent_id] if agent_id else conv
    conv_ids = set(conv_agt["conversation_id"].unique())

    kpi = raw["kpi_results"]
    kpi_agt = kpi[kpi["voiceAgentId"] == agent_id] if agent_id else kpi
    kpi_ids = set(kpi_agt["voiceConversationId"].unique())

    tw = raw["twilio_events"]
    tw_ids = set(tw["conversation_id"].unique())

    rows = []
    for _, row_ in working_table.iterrows():
        cid = row_["conversation_id"]
        in_conv = cid in conv_ids
        in_kpi = cid in kpi_ids
        in_tw = cid in tw_ids

        if in_conv and in_kpi and in_tw:
            status = "All Sources Complete"
        elif in_conv and in_kpi:
            status = "Missing Twilio"
        elif in_conv and in_tw:
            status = "Missing KPI"
        else:
            status = "Missing KPI and Twilio"

        call_date = row_.get("start_dt_pht")
        if pd.notna(call_date):
            call_date = call_date.date()

        rows.append({
            "Conversation ID": cid,
            "In Conversations": "Yes" if in_conv else "No",
            "In KPI Results": "Yes" if in_kpi else "No",
            "In Webhook Events": "Yes" if in_tw else "No",
            "Join Status": status,
            "Agent ID": agent_id,
            "Call Date (PHT)": call_date,
        })

    df = pd.DataFrame(rows)

    # Filter to the report's date range
    if start_date and end_date:
        df = df[
            (df["Call Date (PHT)"] >= start_date)
            & (df["Call Date (PHT)"] <= end_date)
        ]

    return df.sort_values(["Call Date (PHT)", "Conversation ID"]).reset_index(drop=True)


# ── Sheet 2: Field Completeness ───────────────────────────────────

# (Call Detail Log column, label shown when the cell is blank). Every KPI
# field the GFiber agent emits is audited here, so the list is driven off a
# spec rather than a hand-written block per field. "Conversation ID" is the
# row key and is excluded; the value-bearing columns are all that count
# toward the completeness score.
_COMPLETENESS_FIELDS = [
    ("Contact Number", "MISSING"),
    ("Status", "MISSING (No Twilio data)"),
    ("Call Duration (sec)", "MISSING (No call_logs data)"),
    ("Participated Call", "MISSING (No KPI data)"),
    ("Identity Confirmed", "MISSING (No KPI data)"),
    ("Consent & Recording Confirmed", "MISSING (No KPI data)"),
    ("Postpaid Status", "MISSING (No KPI data)"),
    ("Application Intent", "MISSING (No KPI data)"),
    ("Application Completed", "MISSING (No KPI data)"),
    ("Final Disposition", "MISSING (No KPI data)"),
    ("Endorsed for Work Order", "MISSING (No KPI data)"),
    ("Lead for Outbound Handling", "MISSING (No KPI data)"),
    ("Lead for Email Remarketing", "MISSING (No KPI data)"),
    ("Non-Completion Reason", "MISSING (No KPI data)"),
    ("Competitor Detected", "MISSING (No KPI data)"),
    ("Competitor Name", "MISSING (No KPI data)"),
    ("Reason for Switch", "MISSING (No KPI data)"),
    ("Customer Disposition", "MISSING (No KPI data)"),
    ("Customer Sentiment & Feedback", "MISSING (No KPI data)"),
    ("Customer Asked Questions", "MISSING (No KPI data)"),
    ("Question Topics", "MISSING (No KPI data)"),
    ("Customer Questions & Concerns", "MISSING (No KPI data)"),
    ("Opt-Out Flag", "MISSING (No KPI data)"),
    ("Repeat Requested", "MISSING (No KPI data)"),
    ("Identity Re-asked (defect)", "MISSING (No KPI data)"),
    ("Order Number", "MISSING (No KPI data)"),
    ("Call Date (PHT)", "MISSING"),
    ("Call Time (PHT)", "MISSING"),
    ("Call Completed", "MISSING (No KPI data)"),
]

_TOTAL_COMPLETENESS_FIELDS = len(_COMPLETENESS_FIELDS)


def _build_field_completeness(detail_log, start_date=None, end_date=None):
    if start_date and end_date:
        detail_log = detail_log[
            (detail_log["Call Date (PHT)"] >= start_date)
            & (detail_log["Call Date (PHT)"] <= end_date)
        ]

    rows = []

    for _, row_ in detail_log.iterrows():
        entry = {"Conversation ID": row_["Conversation ID"]}
        score = 0

        for column, missing_label in _COMPLETENESS_FIELDS:
            value = row_.get(column)
            if _is_blank(value):
                entry[column] = missing_label
            else:
                # Contact Number and the two timestamp columns are just
                # present-or-not; echoing the value would only duplicate
                # what the Call Detail Log already shows.
                if column in ("Contact Number", "Call Date (PHT)", "Call Time (PHT)"):
                    entry[column] = "Populated"
                elif column == "Call Duration (sec)":
                    entry[column] = f"Populated ({value}s)"
                else:
                    entry[column] = f"Populated ({value})"
                score += 1

        entry["Completeness Score"] = f"{score}/{_TOTAL_COMPLETENESS_FIELDS}"
        rows.append(entry)

    if not rows:
        return pd.DataFrame(columns=(
            ["Conversation ID"]
            + [c for c, _ in _COMPLETENESS_FIELDS]
            + ["Completeness Score"]
        ))

    return pd.DataFrame(rows)


def _is_blank(val):
    """True when a cell is NaN, None, empty-string, or pd.NA."""
    if val is None:
        return True
    if isinstance(val, float) and pd.isna(val):
        return True
    if isinstance(val, str) and val.strip() == "":
        return True
    return False


# ── Sheet 3: Calculation Audit ────────────────────────────────────

def _build_calculation_audit(detail_log, eod_df, start_date, end_date):
    range_log = detail_log[
        (detail_log["Call Date (PHT)"] >= start_date)
        & (detail_log["Call Date (PHT)"] <= end_date)
    ]

    eod_lookup = dict(zip(eod_df["Metric"].astype(str), eod_df["Value"]))

    dialed = len(range_log)
    connected = int((range_log["Status"] == "Connected").sum())
    failed = int((range_log["Status"] == "Failed").sum())
    no_answer = int((range_log["Status"] == "No Answer").sum())
    busy = int((range_log["Status"] == "Busy").sum())
    unmatched = int(range_log["Status"].isna().sum())

    connected_calls = range_log[range_log["Status"] == "Connected"]

    def count(column, value):
        return int((connected_calls[column] == value).sum())

    def count_by_postpaid(column, value, postpaid_value):
        return int((
            (connected_calls[column] == value)
            & (connected_calls["Postpaid Status"] == postpaid_value)
        ).sum())

    contacted = count("Participated Call", "Yes")
    identity_confirmed = count("Identity Confirmed", "Yes")
    wrong_customer = count("Identity Confirmed", "No")
    consented = count("Consent & Recording Confirmed", "Yes")
    declined_recording = count("Consent & Recording Confirmed", "No")
    postpaid = count("Postpaid Status", "postpaid")
    non_postpaid = count("Postpaid Status", "non_postpaid")
    proceed = count("Application Intent", "proceed")
    no_longer_interested = count("Application Intent", "no_longer_interested")
    already_completed = count("Application Intent", "already_completed")
    endorsed = count("Endorsed for Work Order", "Yes")
    lead_outbound = count("Lead for Outbound Handling", "Yes")
    lead_email = count("Lead for Email Remarketing", "Yes")
    competitor_identified = count("Competitor Detected", "Yes")

    proceed_postpaid = count_by_postpaid("Application Intent", "proceed", "postpaid")
    proceed_non_postpaid = count_by_postpaid("Application Intent", "proceed", "non_postpaid")
    nli_postpaid = count_by_postpaid("Application Intent", "no_longer_interested", "postpaid")
    nli_non_postpaid = count_by_postpaid("Application Intent", "no_longer_interested", "non_postpaid")
    completed_postpaid = count_by_postpaid("Application Intent", "already_completed", "postpaid")
    completed_non_postpaid = count_by_postpaid("Application Intent", "already_completed", "non_postpaid")

    conn_rate_val = round((connected / dialed) * 100, 1) if dialed else 0.0
    conv_rate_val = round((proceed / connected) * 100, 1) if connected else 0.0

    def pct_of_connected(n):
        return round((n / connected) * 100, 1) if connected else 0.0

    postpaid_pct = pct_of_connected(postpaid)
    non_postpaid_pct = pct_of_connected(non_postpaid)
    endorsed_pct = pct_of_connected(endorsed)
    lead_outbound_pct = pct_of_connected(lead_outbound)
    lead_email_pct = pct_of_connected(lead_email)

    dur = range_log["Call Duration (sec)"].dropna()
    total_sec = int(dur.sum()) if not dur.empty else 0
    total_min = round(total_sec / 60, 1) if total_sec else None
    avg_dur = round(dur.mean(), 1) if not dur.empty else None

    target_raw = eod_lookup.get("Calls Dialed - Target", "")
    try:
        target_val = float(target_raw)
    except (TypeError, ValueError):
        target_val = 0
    system_errors = max(0, target_val - dialed)

    retries = max(0, failed + no_answer + busy + system_errors)

    rows = []

    _sentinel = object()

    def add_step(step, metric, formula, operands, computed, report_key):
        expected = eod_lookup.get(report_key, _sentinel)
        if expected is _sentinel:
            match = "(not in EOD report)"
            display_expected = "(not in report)"
        else:
            display_expected = expected
            try:
                if str(computed) != str(expected):
                    match = f"MISMATCH (expected {expected})"
                else:
                    match = "PASS"
            except (ValueError, TypeError):
                match = f"MISMATCH (expected {expected})"
        rows.append({
            "Step": step,
            "Metric": metric,
            "Formula / Derivation": formula,
            "Operands / Intermediate Values": operands,
            "Computed Value": computed,
            "EOD Report Value": display_expected,
            "Match": match,
        })

    add_step(1, "Calls Dialed - Actual",
             "COUNT(Call Detail Log rows)", f"{dialed} rows", dialed,
             "Calls Dialed - Actual")
    add_step(2, "Calls Connected",
             "COUNTIF(Status = 'Connected')", f"{connected} rows", connected,
             "Calls Connected")
    add_step(3, "No Answer",
             "COUNTIF(Status = 'No Answer')", f"{no_answer} rows", no_answer,
             "No Answer")
    add_step(4, "Busy",
             "COUNTIF(Status = 'Busy')", f"{busy} rows", busy,
             "Busy")
    add_step(5, "Failed",
             "COUNTIF(Status = 'Failed')", f"{failed} rows", failed,
             "Failed")  # not in report but useful
    add_step(6, "Unmatched (blank Status)",
             "COUNTIF(Status = blank)", f"{unmatched} rows", unmatched,
             "Unmatched")  # not in report
    add_step(7, "Contacted",
             "COUNTIFS(Status='Connected', Participated Call='Yes')",
             f"{connected} connected, {contacted} Yes", contacted,
             "Contacted")
    add_step(8, "Identity Confirmed",
             "COUNTIFS(Status='Connected', Identity Confirmed='Yes')",
             f"{connected} connected, {identity_confirmed} Yes",
             identity_confirmed, "Identity Confirmed")
    add_step(9, "Wrong Customer",
             "COUNTIFS(Status='Connected', Identity Confirmed='No')",
             f"{connected} connected, {wrong_customer} No",
             wrong_customer, "Wrong Customer")
    add_step(10, "Consented to Continue and Recording",
             "COUNTIFS(Status='Connected', Consent & Recording Confirmed='Yes')",
             f"{connected} connected, {consented} Yes",
             consented, "Consented to Continue and Recording")
    add_step(11, "Declined Recording",
             "COUNTIFS(Status='Connected', Consent & Recording Confirmed='No')",
             f"{connected} connected, {declined_recording} No",
             declined_recording, "Declined Recording")
    add_step(12, "Postpaid Verified",
             "COUNTIFS(Status='Connected', Postpaid Status='postpaid')",
             f"{connected} connected, {postpaid} postpaid",
             postpaid, "Postpaid Verified")
    add_step(13, "Non-Postpaid Verified",
             "COUNTIFS(Status='Connected', Postpaid Status='non_postpaid')",
             f"{connected} connected, {non_postpaid} non_postpaid",
             non_postpaid, "Non-Postpaid Verified")
    add_step(14, "Wishes to Proceed",
             "COUNTIFS(Status='Connected', Application Intent='proceed')",
             f"{connected} connected, {proceed} proceed",
             proceed, "Wishes to Proceed")
    add_step(15, "No Longer Interested",
             "COUNTIFS(Status='Connected', Application Intent='no_longer_interested')",
             f"{connected} connected, {no_longer_interested} no_longer_interested",
             no_longer_interested, "No Longer Interested")
    add_step(16, "Application Already Completed",
             "COUNTIFS(Status='Connected', Application Intent='already_completed')",
             f"{connected} connected, {already_completed} already_completed",
             already_completed, "Application Already Completed")
    add_step(17, "Endorsed for Work Order",
             "COUNTIFS(Status='Connected', Endorsed for Work Order='Yes')",
             f"{connected} connected, {endorsed} Yes",
             endorsed, "Endorsed for Work Order")
    add_step(18, "Lead for Outbound Handling",
             "COUNTIFS(Status='Connected', Lead for Outbound Handling='Yes')",
             f"{connected} connected, {lead_outbound} Yes",
             lead_outbound, "Lead for Outbound Handling")
    add_step(19, "Email Remarketing Tagged",
             "COUNTIFS(Status='Connected', Lead for Email Remarketing='Yes')",
             f"{connected} connected, {lead_email} Yes",
             lead_email, "Email Remarketing Tagged")
    add_step(20, "Competitor Identified",
             "COUNTIFS(Status='Connected', Competitor Detected='Yes')",
             f"{connected} connected, {competitor_identified} Yes",
             competitor_identified, "Competitor Identified")
    add_step(21, "Connection Rate",
             "(Connected / Dialed) x 100",
             f"{connected} / {dialed} = {connected/dialed:.4f}" if dialed else "N/A",
             f"{conn_rate_val}%",
             "Connection Rate (Connected / Dialed)")
    add_step(22, "Conversion Rate",
             "(Proceed / Connected) x 100",
             f"{proceed} / {connected} = {proceed/connected:.4f}" if connected else "N/A",
             f"{conv_rate_val}%",
             "Conversion Rate (Proceed / Connected)")
    add_step(23, "Total Call Duration (minutes)",
             "SUM(Call Duration (sec)) / 60",
             f"{total_sec} sec / 60", total_min,
             "Total Call Duration (minutes)")
    add_step(24, "Avg. Call Duration - Connected",
             "AVERAGE(Call Duration (sec) WHERE Status='Connected')",
             f"{len(dur)} rows, sum={total_sec} sec", avg_dur,
             "Avg. Call Duration - Connected (seconds)")
    add_step(25, "System Errors",
             "MAX(0, Calls Dialed - Target - Calls Dialed - Actual)",
             f"{target_val} - {dialed} = {target_val - dialed}", system_errors,
             "System Errors")
    add_step(26, "Retries Queued for Tomorrow",
             "Failed + No Answer + Busy + System Errors",
             f"{failed} + {no_answer} + {busy} + {system_errors} = {retries}",
             retries, "Retries Queued for Tomorrow")

    # Intent x segment cross-tab. These won't sum back to their parent
    # total when some connected calls carry a blank Postpaid Status.
    add_step(27, "Wishes to Proceed - Postpaid",
             "COUNTIFS(Status='Connected', Application Intent='proceed', Postpaid Status='postpaid')",
             f"{proceed} proceed, {proceed_postpaid} postpaid",
             proceed_postpaid, "Wishes to Proceed - Postpaid")
    add_step(28, "Wishes to Proceed - Non-Postpaid",
             "COUNTIFS(Status='Connected', Application Intent='proceed', Postpaid Status='non_postpaid')",
             f"{proceed} proceed, {proceed_non_postpaid} non_postpaid",
             proceed_non_postpaid, "Wishes to Proceed - Non-Postpaid")
    add_step(29, "No Longer Interested - Postpaid",
             "COUNTIFS(Status='Connected', Application Intent='no_longer_interested', Postpaid Status='postpaid')",
             f"{no_longer_interested} no_longer_interested, {nli_postpaid} postpaid",
             nli_postpaid, "No Longer Interested - Postpaid")
    add_step(30, "No Longer Interested - Non-Postpaid",
             "COUNTIFS(Status='Connected', Application Intent='no_longer_interested', Postpaid Status='non_postpaid')",
             f"{no_longer_interested} no_longer_interested, {nli_non_postpaid} non_postpaid",
             nli_non_postpaid, "No Longer Interested - Non-Postpaid")
    add_step(31, "Application Already Completed - Postpaid",
             "COUNTIFS(Status='Connected', Application Intent='already_completed', Postpaid Status='postpaid')",
             f"{already_completed} already_completed, {completed_postpaid} postpaid",
             completed_postpaid, "Application Already Completed - Postpaid")
    add_step(32, "Application Already Completed - Non-Postpaid",
             "COUNTIFS(Status='Connected', Application Intent='already_completed', Postpaid Status='non_postpaid')",
             f"{already_completed} already_completed, {completed_non_postpaid} non_postpaid",
             completed_non_postpaid, "Application Already Completed - Non-Postpaid")

    # Share-of-connected percentages. Computed values carry the "%" suffix
    # so they compare equal to the string the EOD Report writes.
    add_step(33, "Postpaid Customers %",
             "(Postpaid Verified / Connected) x 100",
             f"{postpaid} / {connected}" if connected else "N/A",
             f"{postpaid_pct}%", "Postpaid Customers % (of Connected)")
    add_step(34, "Non-Postpaid Customers %",
             "(Non-Postpaid Verified / Connected) x 100",
             f"{non_postpaid} / {connected}" if connected else "N/A",
             f"{non_postpaid_pct}%", "Non-Postpaid Customers % (of Connected)")
    add_step(35, "Endorsed for Work Order %",
             "(Endorsed for Work Order / Connected) x 100",
             f"{endorsed} / {connected}" if connected else "N/A",
             f"{endorsed_pct}%", "Endorsed for Work Order % (of Connected)")
    add_step(36, "Lead for Outbound Handling %",
             "(Lead for Outbound Handling / Connected) x 100",
             f"{lead_outbound} / {connected}" if connected else "N/A",
             f"{lead_outbound_pct}%", "Lead for Outbound Handling % (of Connected)")
    add_step(37, "Email Remarketing %",
             "(Email Remarketing Tagged / Connected) x 100",
             f"{lead_email} / {connected}" if connected else "N/A",
             f"{lead_email_pct}%", "Email Remarketing % (of Connected)")

    return pd.DataFrame(rows)


# ── Sheet 4: Data Quality Issues ──────────────────────────────────

def _build_data_quality_issues(working_table, detail_log,
                                start_date=None, end_date=None):
    if start_date and end_date:
        detail_log = detail_log[
            (detail_log["Call Date (PHT)"] >= start_date)
            & (detail_log["Call Date (PHT)"] <= end_date)
        ]
        keep_ids = set(detail_log["Conversation ID"])
        working_table = working_table[working_table["conversation_id"].isin(keep_ids)]

    issues = []

    # Issues from the detail log rows
    for _, row_ in detail_log.iterrows():
        cid = row_["Conversation ID"]

        if _is_blank(row_.get("Status")):
            issues.append({
                "Conversation ID": cid,
                "Issue": "Missing Twilio Event Data",
                "Detail": (
                    "No matching twilio_webhook_events found for this "
                    "conversation. Status is blank."
                ),
                "Severity": "Medium",
            })

        # Check for missing KPI data by looking at Customer Disposition
        # (Application Completed is never blank; it returns "N/A" when KPI
        # data is absent, so it can't serve as the missingness signal.)
        if _is_blank(row_.get("Customer Disposition")):
            issues.append({
                "Conversation ID": cid,
                "Issue": "Missing KPI Results Data",
                "Detail": (
                    "No matching kpi_results found for this conversation. "
                    "All KPI-derived fields (Application Intent, Final "
                    "Disposition, Customer Disposition, etc.) are blank; "
                    "Application Completed is 'N/A'."
                ),
                "Severity": "High",
            })

        if _is_blank(row_.get("Call Duration (sec)")):
            issues.append({
                "Conversation ID": cid,
                "Issue": "Missing Call Duration",
                "Detail": (
                    "call_logs field is null or in an unparseable format. "
                    "Duration cannot be extracted."
                ),
                "Severity": "Medium",
            })

    # Contact-number reliability issues from the working table
    for _, row_ in working_table.iterrows():
        cid = row_["conversation_id"]
        reliability = row_.get("contact_number_reliability")
        if pd.notna(reliability) and "TRUNCATED" in str(reliability):
            issues.append({
                "Conversation ID": cid,
                "Issue": "Truncated Contact Number",
                "Detail": str(reliability),
                "Severity": "Medium",
            })

        if pd.notna(reliability) and (
            "Unparseable" in str(reliability) or "unexpected length" in str(reliability)
        ):
            issues.append({
                "Conversation ID": cid,
                "Issue": "Invalid Contact Number Format",
                "Detail": str(reliability),
                "Severity": "High",
            })

    if not issues:
        return pd.DataFrame(columns=[
            "Conversation ID", "Issue", "Detail", "Severity",
        ])

    return pd.DataFrame(issues)


# ── Sheet 5: Duplicate Contacts ───────────────────────────────────

def _build_duplicate_contacts(working_table, start_date=None, end_date=None):
    """
    Audit trail for same-day duplicate-number dedup: every row that took
    part in a (Contact Number, Call Date) collision, which one dedup kept,
    and how many rows were in the group. Cross-format duplicates (e.g.
    "639673187061" vs "09673187061") are caught here too, since both the
    raw log and the final log key off the same normalized Contact Number
    (see preprocessing.normalize_ph_digits).
    """
    columns = [
        "Contact Number", "Call Date (PHT)", "Conversation ID",
        "Call Time (PHT)", "Status", "Outcome", "Duplicate Group Size",
    ]

    raw_log = call_detail.build_raw_call_rows(working_table)
    if start_date and end_date:
        raw_log = raw_log[
            (raw_log["Call Date (PHT)"] >= start_date)
            & (raw_log["Call Date (PHT)"] <= end_date)
        ]

    has_number = raw_log["Contact Number"].notna()
    dupable = raw_log[has_number].copy()
    if dupable.empty:
        return pd.DataFrame(columns=columns)

    dupable["_group_size"] = dupable.groupby(
        ["Contact Number", "Call Date (PHT)"]
    )["Conversation ID"].transform("size")
    dup_rows = dupable[dupable["_group_size"] > 1].copy()

    if dup_rows.empty:
        return pd.DataFrame(columns=columns)

    kept_ids = set(call_detail.build_call_detail_log(working_table)["Conversation ID"])
    dup_rows["Outcome"] = dup_rows["Conversation ID"].apply(
        lambda cid: "Kept" if cid in kept_ids else "Merged (duplicate)"
    )
    dup_rows = dup_rows.rename(columns={"_group_size": "Duplicate Group Size"})

    return dup_rows.sort_values(
        ["Contact Number", "Call Date (PHT)", "Call Time (PHT)"]
    )[columns].reset_index(drop=True)


# ── Public entry point ────────────────────────────────────────────

def build_validation_report(working_table, detail_log, eod_df,
                             start_date, end_date, agent_id):
    """
    Build all 5 validation-report sheets and return them as a
    {sheet_key: DataFrame} dict ready for write_validation_report().
    """
    return {
        "join_summary": _build_join_summary(working_table, agent_id,
                                            start_date, end_date),
        "field_completeness": _build_field_completeness(detail_log,
                                                        start_date, end_date),
        "calculation_audit": _build_calculation_audit(detail_log, eod_df,
                                                      start_date, end_date),
        "data_quality_issues": _build_data_quality_issues(working_table,
                                                          detail_log,
                                                          start_date, end_date),
        "duplicate_contacts": _build_duplicate_contacts(working_table,
                                                        start_date, end_date),
    }
