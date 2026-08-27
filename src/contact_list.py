"""
contact_list.py
------------------
Business logic for Mode 2: the GFiber Application Contact List. Takes the
raw customer_phone / user list Globe provides and normalizes each number to
the +63XXXXXXXXXX form used for dialling.

Also handles data validation and categorization for the validation report.
"""

import re
from collections import Counter
import pandas as pd


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


# ── Categorization ─────────────────────────────────────────────────


def categorize_records(raw_df: pd.DataFrame) -> dict:
    """
    Validates and categorizes every row in the raw contact list.

    Processing order per row:
      1. Phone chain: missing → code → length (stops on first failure)
      2. Name check (independent)
      3. Duplicate check (global, always runs)

    Returns a dict with two DataFrames:
        valid     — passes all checks, in input file order
        invalid   — failed at least one check (+ ``reason`` column)
    """
    df = raw_df.copy()

    # Phase 1: validate each row
    processed = []
    for _, row in df.iterrows():
        phone_raw = row.get("customer_phone")
        name_raw = row.get("user")

        reasons = []
        phone_display = str(phone_raw).strip() if not pd.isna(phone_raw) else ""
        name_display = str(name_raw).strip() if not pd.isna(name_raw) else ""

        # Phone chain (stops on first failure)
        phone_ok, phone_reason, phone_9x, _ = validate_phone_format(phone_raw)
        if phone_ok:
            normalized_phone = "+63" + phone_9x
        else:
            normalized_phone = phone_display
            reasons.append(phone_reason)

        # Name check (independent of the other chains)
        name_ok, name_reason, clean_name = validate_customer_name(name_raw)
        if not name_ok:
            reasons.append(name_reason)

        processed.append({
            "phone_raw": phone_display,
            "normalized_phone": normalized_phone,
            "name_raw": name_display,
            "clean_name": clean_name,
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
            invalid_list.append({
                "customer_phone": p["phone_raw"],
                "user": p["name_raw"],
                "reason": reason_str,
            })
        else:
            valid_list.append({
                "customer_phone": p["normalized_phone"],
                "user": p["clean_name"],
            })

    valid_cols = ["customer_phone", "user"]
    invalid_cols = ["customer_phone", "user", "reason"]

    # No priority signal to sort on — records keep their input file order.
    return {
        "valid": pd.DataFrame(valid_list, columns=valid_cols),
        "invalid": pd.DataFrame(invalid_list, columns=invalid_cols),
    }
