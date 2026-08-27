"""
call_detail.py
----------------
Builds the "Call Detail Log" sheet: one row per individual call, carrying
every field the GFiber abandoned-application agent emits in outputJson
alongside the Twilio call-progress outcome.
"""

import pandas as pd

# Priority tiers for picking the "best" record among same-day duplicate
# dials to the same number. "Connected" always wins outright. Busy/No
# Answer/Failed are no longer ranked against each other - among those (or
# any other non-blank status), the latest Call Time wins instead. A blank/
# unmatched status is always lowest priority.
_STATUS_PRIORITY = {"Connected": 0}
_NON_BLANK_STATUS_PRIORITY = 1
_BLANK_STATUS_PRIORITY = 2


def _blank_if_missing(value):
    """Turns NaN into a real Python None so openpyxl writes a blank cell
    instead of the literal string 'nan'."""
    return None if pd.isna(value) else value


def _yes_no(value):
    """Yes/No for a KPI-derived boolean flag; blank if missing."""
    if pd.isna(value):
        return None
    return "Yes" if bool(value) else "No"


def _yes_no_na(value):
    """
    Yes/No/N/A for a KPI-derived boolean flag. The agent emits the literal
    string "NA" for not-applicable booleans (seen in real kpi_results
    exports), so that is treated the same as a missing value.
    """
    if pd.isna(value) or (isinstance(value, str) and value.strip().upper() == "NA"):
        return "N/A"
    return "Yes" if bool(value) else "No"


def _format_list(value):
    """Joins a list of strings into a comma-separated display string.
    Empty lists and missing values both render as a blank cell."""
    if not isinstance(value, list) or not value:
        return None
    return ", ".join(str(v) for v in value)


def _map_status(twilio_status):
    """
    Maps Twilio call stages to display-friendly status labels.

    Mapping:
        completed, in-progress -> Connected
        no-answer -> No Answer
        busy -> Busy
        failed -> Failed
        ringing -> No Answer
    """
    if pd.isna(twilio_status):
        return None

    status_map = {
        "completed": "Connected",
        "in-progress": "Connected",
        "no-answer": "No Answer",
        "busy": "Busy",
        "failed": "Failed",
        "ringing": "No Answer",
    }

    return status_map.get(twilio_status, twilio_status)


def _dedupe_duplicate_calls(log: pd.DataFrame) -> pd.DataFrame:
    """
    Collapses repeat same-day dials to the same Contact Number down to one
    record per (Contact Number, Call Date (PHT)):

      - "Connected" always wins outright if present among the duplicates.
      - Otherwise (including ties among Busy/No Answer/Failed, or repeats
        of the same status), keep the latest record by Call Time.

    Rows with a blank Contact Number are never collapsed against each
    other (no reliable way to confirm they're the same customer).
    """
    has_number = log["Contact Number"].notna()
    dedupable = log[has_number].copy()
    passthrough = log[~has_number]

    if dedupable.empty:
        return log

    dedupable["_priority"] = dedupable["Status"].apply(
        lambda s: _BLANK_STATUS_PRIORITY if pd.isna(s)
        else _STATUS_PRIORITY.get(s, _NON_BLANK_STATUS_PRIORITY)
    )
    dedupable = dedupable.sort_values(
        ["Contact Number", "Call Date (PHT)", "_priority", "Call Time (PHT)"],
        ascending=[True, True, True, False],
    )
    deduped = dedupable.drop_duplicates(
        subset=["Contact Number", "Call Date (PHT)"], keep="first"
    ).drop(columns="_priority")

    return pd.concat([deduped, passthrough], ignore_index=True)


def _kpi(df: pd.DataFrame, field: str) -> pd.Series:
    """
    Reads a KPI column out of the working table, tolerating its absence.
    outputJson payloads vary — a field no conversation populated never
    becomes a column at all, so a plain df[field] would raise KeyError.
    """
    return df.get(field, pd.Series(dtype=object, index=df.index))


def build_raw_call_rows(working_table: pd.DataFrame) -> pd.DataFrame:
    """
    Transforms the merged working table into one row per individual call,
    before same-day duplicate-number collapsing. Exposed separately from
    build_call_detail_log so callers (e.g. the validation report) can audit
    which rows got collapsed as duplicates.

    Status: sourced exclusively from the Twilio call-progress journey
    (twilio_final_status, derived in preprocessing.extract_twilio_details
    from twilio_webhook_events.csv). Mapped to display-friendly values:
    "Connected", "No Answer", "Busy", "Failed". If a conversation_id has no
    matching Twilio events, Status is left blank.
    """
    df = working_table.copy()

    return pd.DataFrame({
        "Conversation ID": df["conversation_id"],
        "Contact Number": df["contact_number_clean"],
        "Status": df["twilio_final_status"].apply(_map_status).apply(_blank_if_missing),
        "Call Duration (sec)": df["call_duration_sec"],
        "Participated Call": _kpi(df, "participated_call").apply(_yes_no),
        "Identity Confirmed": _kpi(df, "identity_confirmed").apply(_yes_no),
        "Consent & Recording Confirmed": _kpi(df, "consent_and_recording_confirmed").apply(_yes_no),
        "Postpaid Status": _kpi(df, "is_postpaid_customer"),
        "Application Intent": _kpi(df, "application_intent"),
        "Application Completed": _kpi(df, "application_completed").apply(_yes_no_na),
        "Final Disposition": _kpi(df, "final_disposition"),
        "Endorsed for Work Order": _kpi(df, "endorsed_for_work_order").apply(_yes_no),
        "Lead for Outbound Handling": _kpi(df, "lead_for_outbound_handling").apply(_yes_no),
        "Lead for Email Remarketing": _kpi(df, "lead_for_email_remarketing").apply(_yes_no),
        "Non-Completion Reason": _kpi(df, "non_completion_reason"),
        "Competitor Detected": _kpi(df, "competitor_detected").apply(_yes_no),
        "Competitor Name": _kpi(df, "competitor_name"),
        "Reason for Switch": _kpi(df, "reason_for_switch"),
        "Customer Disposition": _kpi(df, "customer_disposition"),
        "Customer Sentiment & Feedback": _kpi(df, "customer_sentiment_feedback"),
        "Customer Asked Questions": _kpi(df, "customer_asked_questions").apply(_yes_no),
        "Question Topics": _kpi(df, "question_topics").apply(_format_list),
        "Customer Questions & Concerns": _kpi(df, "customer_questions_concerns"),
        "Opt-Out Flag": _kpi(df, "opt_out_flag").apply(_yes_no),
        "Repeat Requested": _kpi(df, "repeat_requested").apply(_yes_no),
        "Identity Re-asked (defect)": _kpi(df, "identity_reconfirmation_requested").apply(_yes_no),
        "Order Number": _kpi(df, "order_number"),
        "Call Date (PHT)": df["start_dt_pht"].dt.date,
        "Call Time (PHT)": df["start_dt_pht"].dt.strftime("%H:%M:%S"),
        "Call Completed": _kpi(df, "call_completed").apply(_yes_no),
    })


def build_call_detail_log(working_table: pd.DataFrame) -> pd.DataFrame:
    """
    Transforms the merged working table into the final Call Detail Log
    DataFrame, ready to write to Excel (one row per Contact Number per
    Call Date - see _dedupe_duplicate_calls).
    """
    log = build_raw_call_rows(working_table)
    log = _dedupe_duplicate_calls(log)
    return log.sort_values(["Call Date (PHT)", "Call Time (PHT)"]).reset_index(drop=True)
