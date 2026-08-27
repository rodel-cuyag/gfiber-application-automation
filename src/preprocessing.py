"""
preprocessing.py
------------------
Cleans and joins the 3 raw tables into a single "working" DataFrame,
one row per conversation, filtered to the target agent_id.

Join logic:
  conversations.conversation_id  <-> kpi_results.voiceConversationId
  conversations.conversation_id  <-> twilio_webhook_events.conversation_id
  agent filter uses conversations.agent_id AND kpi_results.voiceAgentId
"""

import json
import re
import pandas as pd
from src import config


def _safe_json_loads(value):
    """Parse a JSON string safely; return {} if null/invalid/empty."""
    if pd.isna(value) or value in ("", "{}"):
        return {}
    try:
        return json.loads(value)
    except (json.JSONDecodeError, TypeError):
        return {}


def filter_conversations_by_agent(conversations: pd.DataFrame, agent_id: int) -> pd.DataFrame:
    """Keep only conversations belonging to the target agent_id."""
    return conversations[conversations["agent_id"] == agent_id].copy()


def add_local_timestamps(conversations: pd.DataFrame) -> pd.DataFrame:
    """
    Convert epoch-millisecond UTC timestamps into PHT datetime columns.
    Adds: start_dt_pht, end_dt_pht, call_date (date-only, used for EOD
    grouping/date-range filtering), call_duration_sec.

    start_timestamp is confirmed reliable — converts to a real, sane PHT
    datetime and is what we use to bucket calls into a Call Date / the
    --start-date/--end-date range.

    end_timestamp converts to a technically-valid datetime too (the
    epoch-ms math is fine), but the *values themselves* aren't
    trustworthy as a real call-end moment: for this dataset, end_timestamp
    is frequently identical to or earlier than start_timestamp (checked
    across 1,402 rows with both fields populated — ~79% have zero diff,
    ~20% have a *negative* diff of -20 to -30 million ms, and it doesn't
    correlate at all with the real call length in call_logs). We still
    compute end_dt_pht for visibility/audit purposes, but duration is
    deliberately NOT derived from it — see extract_call_duration().
    """
    df = conversations.copy()
    df["start_dt_pht"] = (
        pd.to_datetime(df["start_timestamp"], unit="ms", utc=True)
        .dt.tz_convert(config.TIMEZONE)
    )
    df["end_dt_pht"] = (
        pd.to_datetime(df["end_timestamp"], unit="ms", utc=True)
        .dt.tz_convert(config.TIMEZONE)
    )
    df["call_date"] = df["start_dt_pht"].dt.date
    df["call_duration_sec"] = extract_call_duration(df)
    return df


def extract_call_duration(conversations: pd.DataFrame) -> pd.Series:
    """
    Derives call duration (seconds) from call_logs.metrics.total_duration_ms.

    NOTE: end_timestamp in conversations.csv is unreliable for this purpose —
    many rows share identical/placeholder end_timestamp values that land
    *before* start_timestamp, producing nonsense negative durations. The
    per-call metrics embedded in call_logs are the trustworthy source, so
    we use that instead and ignore end_timestamp entirely.
    """
    def duration_seconds(call_logs_json):
        logs = _safe_json_loads(call_logs_json)
        # Some agents store call_logs as a list of turn-by-turn bot/user
        # events rather than the {"metrics": {...}} dict shape agent 1060
        # uses. Not this task's scope to parse that schema for duration —
        # bail out to None (shown as blank) rather than crashing or guessing.
        if not isinstance(logs, dict):
            return None
        ms = logs.get("metrics", {}).get("total_duration_ms")
        return round(ms / 1000) if ms is not None else None

    return conversations["call_logs"].apply(duration_seconds)


def normalize_ph_digits(digits: str):
    """
    Normalizes a bare digit string to the 63XXXXXXXXXX PH mobile format
    where possible. Returns (normalized, is_valid); is_valid is False when
    the digit count/prefix doesn't match a recognized PH mobile pattern
    (normalized still holds the digits as-is in that case, or None if
    there were no digits at all). Shared by clean_contact_number and the
    Twilio "To" number recovery in extract_twilio_details so numbers from
    either source always land in the same canonical form (e.g.
    "639673187061", "09673187061", and "9673187061" all normalize to the
    same "639673187061").
    """
    if len(digits) == 11 and digits.startswith("0"):
        return "63" + digits[1:], True
    if len(digits) == 10 and not digits.startswith("0"):
        return "63" + digits, True
    if len(digits) == 12 and digits.startswith("63"):
        return digits, True
    if digits:
        return digits, False
    return None, False


def clean_contact_number(raw_value):
    """
    Cleans a single contact_number value and reports how trustworthy the
    result is. Returns (cleaned_number, reliability) where reliability is
    a short human-readable label — never silently presented as complete
    when it isn't.

    Two corruption patterns show up in the source data:

    1. Excel scientific notation, e.g. "6.39151E+11" (~95% of rows).
       This happens *upstream*, before the CSV is exported — Excel only
       keeps 5-6 significant digits in that notation, so the true
       trailing digits of the phone number are permanently gone. We
       reconstruct the zero-padded integer so it at least displays
       cleanly (no more "E+11"), but we flag it as TRUNCATED rather than
       pretending the padded zeros are real digits.

    2. Plain, uncorrupted digit strings, e.g. "09176881179" or
       "9151427721" (~4.5% of rows). These are genuinely complete — we
       just normalize them to the full 63XXXXXXXXXX format.
    """
    if pd.isna(raw_value):
        return None, "Missing"

    v = str(raw_value).strip()

    if "E+" in v.upper():
        mantissa = v.upper().split("E+")[0]
        sig_figs = len(mantissa.replace(".", "").replace("-", ""))
        try:
            full_digits = str(int(float(v)))
        except (ValueError, TypeError):
            return v, "Unparseable"
        return (
            full_digits,
            f"TRUNCATED - only first {sig_figs} digits are real, "
            f"rest lost upstream (Excel scientific-notation export)",
        )

    digits = re.sub(r"\D", "", v)
    if digits:
        normalized, is_valid = normalize_ph_digits(digits)
        if is_valid:
            return normalized, "Complete"
        return normalized, "Complete - unexpected length, verify manually"
    return v, "Unparseable"


def add_clean_contact_numbers(conversations: pd.DataFrame) -> pd.DataFrame:
    """Adds contact_number_clean / contact_number_reliability columns."""
    df = conversations.copy()
    if df.empty:
        df["contact_number_clean"] = pd.Series(dtype=object)
        df["contact_number_reliability"] = pd.Series(dtype=object)
        return df
    cleaned = df["contact_number"].apply(clean_contact_number)
    df["contact_number_clean"] = cleaned.apply(lambda t: t[0])
    df["contact_number_reliability"] = cleaned.apply(lambda t: t[1])
    return df


def extract_kpi_fields(kpi_results: pd.DataFrame, agent_id: int) -> pd.DataFrame:
    """
    Filter KPI rows to the target agent (voiceAgentId) and flatten the
    outputJson blob into real columns we can merge on conversation_id.
    """
    kpi = kpi_results[kpi_results["voiceAgentId"] == agent_id].copy()
    parsed = kpi["outputJson"].apply(_safe_json_loads).apply(pd.Series)

    flat = pd.concat(
        [kpi[["voiceConversationId"]].reset_index(drop=True), parsed.reset_index(drop=True)],
        axis=1,
    )
    flat = flat.rename(columns={"voiceConversationId": "conversation_id"})

    # Only keep the fields relevant to the GFiber abandoned-application use
    # case — this is the full set of conversation_metrics the agent defines.
    # (Other agent types may populate different keys in outputJson —
    # missing columns are simply absent after this filter, handled downstream.)
    keep_cols = [
        "conversation_id",
        "participated_call", "identity_confirmed", "consent_and_recording_confirmed",
        "is_postpaid_customer", "application_intent", "application_completed",
        "final_disposition", "endorsed_for_work_order",
        "lead_for_outbound_handling", "lead_for_email_remarketing",
        "non_completion_reason", "competitor_detected", "competitor_name",
        "reason_for_switch", "customer_disposition", "customer_sentiment_feedback",
        "customer_asked_questions", "question_topics", "customer_questions_concerns",
        "opt_out_flag", "repeat_requested", "identity_reconfirmation_requested",
        "order_number", "call_completed",
    ]
    existing_cols = [c for c in keep_cols if c in flat.columns]
    return flat[existing_cols]


def extract_twilio_details(twilio_events: pd.DataFrame) -> pd.DataFrame:
    """
    Flattens the Twilio 'event' JSON blob per conversation_id into three
    derived fields:

      - twilio_final_status: the terminal outcome of the call's Twilio
        lifecycle (completed / no-answer / busy / failed). This is now
        the ONLY source for the Call Detail Log's Status column — there
        is no fallback to conversations.status. If a conversation_id
        never appears in twilio_webhook_events.csv, twilio_final_status
        stays null and the Status column is left blank downstream.

      - twilio_contact_number: the clean, complete "To" number Twilio
        recorded for the call. Used to backfill contact numbers that
        arrived corrupted from conversations.csv, wherever a Twilio
        match happens to exist.

      - twilio_latest_sequence: the highest sequence number reached in
        the call journey (0=initiated, 1=ringing, 2=busy/no-answer/in-progress,
        3=completed). Indicates how far the call progressed.

    NOTE: twilio_webhook_events.csv only has 128 rows and covers a
    handful of agents — for agents with zero matching conversation_ids,
    all three derived fields are simply blank for every row of that agent.
    That's expected, not a bug.
    """
    if twilio_events.empty:
        return pd.DataFrame(columns=["conversation_id", "twilio_final_status", "twilio_contact_number", "twilio_latest_sequence"])

    def parse_row(event_json: str):
        events = _safe_json_loads(event_json)
        if not events:
            return pd.Series({
                "twilio_final_status": None,
                "twilio_contact_number": None,
                "twilio_latest_sequence": None
            })

        # Terminal-status priority: only one of these ever appears
        # alongside the in-flight stages (ringing/initiated/in-progress)
        # in this data, so checking in this order reliably picks the
        # actual outcome of the call.
        priority = ["completed", "no-answer", "busy", "failed", "in-progress", "ringing"]
        final_status = next((status for status in priority if status in events), None)

        # Find the latest sequence number (highest number = furthest stage reached)
        latest_sequence = None
        contact_number = None

        for stage_details in events.values():
            if isinstance(stage_details, dict):
                # Extract contact number, normalized to the same
                # 63XXXXXXXXXX form as contact_number_clean so numbers
                # recovered from Twilio match up with conversations.csv
                # numbers for duplicate detection.
                if stage_details.get("To") and contact_number is None:
                    to_digits = re.sub(r"\D", "", stage_details["To"])
                    contact_number, _ = normalize_ph_digits(to_digits)

                # Track highest sequence number
                seq = stage_details.get("SequenceNumber")
                if seq is not None:
                    try:
                        seq_num = int(seq)
                        if latest_sequence is None or seq_num > latest_sequence:
                            latest_sequence = seq_num
                    except (ValueError, TypeError):
                        pass

        return pd.Series({
            "twilio_final_status": final_status,
            "twilio_contact_number": contact_number,
            "twilio_latest_sequence": latest_sequence
        })

    parsed = twilio_events["event"].apply(parse_row)
    result = pd.concat([twilio_events[["conversation_id"]].reset_index(drop=True), parsed.reset_index(drop=True)], axis=1)
    return result.drop_duplicates(subset="conversation_id")


def build_working_table(agent_id: int) -> pd.DataFrame:
    """
    Main entry point: loads all 3 CSVs, filters to agent_id, and merges
    them into one row-per-conversation DataFrame ready for reporting.
    """
    from src import data_loader  # local import avoids circular import

    raw = data_loader.load_all()

    conversations = filter_conversations_by_agent(raw["conversations"], agent_id)
    conversations = add_local_timestamps(conversations)
    conversations = add_clean_contact_numbers(conversations)

    kpi_flat = extract_kpi_fields(raw["kpi_results"], agent_id)
    twilio_flat = extract_twilio_details(raw["twilio_events"])

    merged = conversations.merge(kpi_flat, on="conversation_id", how="left")
    merged = merged.merge(twilio_flat, on="conversation_id", how="left")

    # Backfill contact number from Twilio's "To" field wherever a match
    # exists — this is a genuinely complete number, not a guess, so it
    # overrides the (possibly truncated) conversations.csv value.
    has_twilio_number = merged["twilio_contact_number"].notna()
    merged.loc[has_twilio_number, "contact_number_clean"] = merged.loc[has_twilio_number, "twilio_contact_number"]
    merged.loc[has_twilio_number, "contact_number_reliability"] = "Complete (recovered from Twilio)"

    return merged