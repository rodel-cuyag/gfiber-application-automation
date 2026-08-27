"""
contact_list.py
------------------
Business logic for Mode 2: the GFiber Application Contact List. Takes the
raw customer_phone / user / application_date list Globe provides and
derives the phone number variants needed for scheduling, plus the
application aging used to prioritise who gets called first.

Also handles data validation and categorization for the validation report.
"""

import re
from collections import Counter
import pandas as pd


def compute_days_since_applied(application_date, as_of_date):
    """
    Plain calendar-day difference (e.g. today Aug 26, application_date
    Aug 10 -> 16). Negative if the application date is in the future,
    which the caller treats as a data error.
    """
    if pd.isna(application_date):
        return None
    applied = application_date.date() if hasattr(application_date, "date") else application_date
    return (as_of_date - applied).days


# ── Validation helpers ─────────────────────────────────────────────


def validate_phone_format(raw_value):
    """
    Validate phone number format. Accepts +63/63 (12 digits), 09 (11 digits),
    and 9 (10 digits) formats, after stripping all non-digit characters.

    Returns (is_valid, reason, phone_9x, last_4).
    """
    if pd.isna(raw_value) or str(raw_value).strip() == "":
        return False, "Missing phone number", None, None

    digits = re.sub(r"\D", "", str(raw_value))

    if digits.startswith("63"):
        if len(digits) != 12:
            return False, "Invalid PH number length / Invalid last 4 digits", None, None
        phone_9x = digits[2:]
    elif digits.startswith("09"):
        if len(digits) != 11:
            return False, "Invalid PH number length / Invalid last 4 digits", None, None
        phone_9x = digits[1:]
    elif digits.startswith("9"):
        if len(digits) != 10:
            return False, "Invalid PH number length / Invalid last 4 digits", None, None
        phone_9x = digits
    else:
        return False, "Invalid PH code (must start with +63, 63, 09, or 9)", None, None

    last_4 = phone_9x[-4:]
    return True, None, phone_9x, last_4


def validate_customer_name(raw_value):
    """
    The agent's opening message and every re-verify/close spiel interpolate
    {user}, so a blank name would put a broken sentence on the call. Treat
    it as invalid rather than dialling with a hole in the script.

    Returns (is_valid, reason, cleaned_name).
    """
    if pd.isna(raw_value) or str(raw_value).strip() == "":
        return False, "Missing customer name", None
    return True, None, str(raw_value).strip()


def validate_application_date(raw_value, as_of_date):
    """
    Validate the abandoned application's date.
    Returns (is_valid, reason, parsed_date, days_since_applied).
    """
    if pd.isna(raw_value) or str(raw_value).strip() == "":
        return False, "Missing application date", None, None

    try:
        parsed = pd.to_datetime(raw_value)
        parsed_date = parsed.date() if hasattr(parsed, "date") else parsed
        days_since = compute_days_since_applied(parsed_date, as_of_date)
        return True, None, parsed_date, days_since
    except (ValueError, TypeError):
        return False, "Invalid date format", None, None


# ── Categorization ─────────────────────────────────────────────────


def categorize_records(raw_df: pd.DataFrame, as_of_date) -> dict:
    """
    Validates and categorizes every row in the raw contact list.

    Processing order per row:
      1. Phone chain: missing → code → length (stops on first failure)
      2. Name check (independent)
      3. Date chain: missing → format (independent)
      4. Duplicate check (global, always runs)

    Returns a dict with two DataFrames:
        valid     — passes all checks, sorted oldest application first
        invalid   — failed at least one check (+ ``reason`` column)
    """
    df = raw_df.copy()

    # Phase 1: validate each row
    processed = []
    for _, row in df.iterrows():
        phone_raw = row.get("customer_phone")
        name_raw = row.get("user")
        applied_raw = row.get("application_date")

        reasons = []
        phone_display = str(phone_raw).strip() if not pd.isna(phone_raw) else ""
        name_display = str(name_raw).strip() if not pd.isna(name_raw) else ""

        # Phone chain (stops on first failure)
        phone_ok, phone_reason, phone_9x, last_4 = validate_phone_format(phone_raw)
        if phone_ok:
            normalized_phone = "+63" + phone_9x
        else:
            normalized_phone = phone_display
            reasons.append(phone_reason)

        # Name check (independent of the other chains)
        name_ok, name_reason, clean_name = validate_customer_name(name_raw)
        if not name_ok:
            reasons.append(name_reason)

        # Date chain (independent of the other chains)
        date_ok, date_reason, parsed_date, days_since = validate_application_date(
            applied_raw, as_of_date
        )
        if not date_ok:
            reasons.append(date_reason)

        processed.append({
            "phone_raw": phone_display,
            "normalized_phone": normalized_phone,
            "name_raw": name_display,
            "clean_name": clean_name,
            "applied_raw": applied_raw,
            "parsed_date": parsed_date,
            "phone_9x": phone_9x,
            "last_4": last_4,
            "days_since_applied": days_since,
            "reasons": reasons,
        })

    # Phase 2: duplicate detection (global, always runs)
    phone_counts = Counter(p["phone_raw"] for p in processed if p["phone_raw"])
    for p in processed:
        if p["phone_raw"] and phone_counts[p["phone_raw"]] > 1:
            p["reasons"].append("Duplicate phone number")

    # Phase 3: classify into two buckets
    valid_list, invalid_list = [], []

    for p in processed:
        reason_str = "; ".join(p["reasons"]) if p["reasons"] else None

        if reason_str:
            # Show parsed date if available, otherwise original raw value
            if p["parsed_date"] is not None:
                applied_display = p["parsed_date"]
            elif not pd.isna(p["applied_raw"]):
                applied_display = str(p["applied_raw"])
            else:
                applied_display = ""
            invalid_list.append({
                "customer_phone": p["phone_raw"],
                "user": p["name_raw"],
                "application_date": applied_display,
                "reason": reason_str,
            })
        else:
            valid_list.append({
                "customer_phone": p["normalized_phone"],
                "customer_phone_9x": p["phone_9x"],
                "last_four_digits": p["last_4"],
                "user": p["clean_name"],
                "application_date": p["parsed_date"],
                "days_since_applied": p["days_since_applied"],
            })

    valid_cols = [
        "customer_phone", "customer_phone_9x", "last_four_digits",
        "user", "application_date", "days_since_applied",
    ]
    invalid_cols = ["customer_phone", "user", "application_date", "reason"]

    if valid_list:
        # Oldest abandoned application first — the aging is the priority
        # signal, so the longest-neglected applications get dialled first.
        valid_df = (
            pd.DataFrame(valid_list)
            .sort_values("days_since_applied", ascending=False)
            .reset_index(drop=True)
        )
    else:
        valid_df = pd.DataFrame(columns=valid_cols)

    return {
        "valid": valid_df,
        "invalid": pd.DataFrame(invalid_list, columns=invalid_cols) if invalid_list else pd.DataFrame(columns=invalid_cols),
    }
