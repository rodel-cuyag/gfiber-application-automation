"""
data_loader.py
--------------
Responsible for ONE thing: reading the raw input files off disk into
pandas DataFrames. No cleaning, no merging, no business logic here —
that lives in preprocessing.py (EOD mode) / contact_list.py (Contact
List mode). Keeping this separate makes it easy to swap files for a
database or API later without touching anything else.

For EOD mode, instead of requiring hardcoded filenames, the loader scans
the data/eod/ directory for CSV files and identifies each one by its
column headers (signature matching). Files can be named anything —
"kpi_results 2.csv", "twilio webhook events.csv", etc. — as long as
the required columns are present.
"""

from pathlib import Path

import pandas as pd
from src import config
from src.progress import Spinner


class MissingInputFileError(Exception):
    """Raised when a required input file is not found on disk."""
    pass


class MissingHeaderError(Exception):
    """Raised when the contact list input is missing required columns."""
    pass


# ── Mode 1: EOD Report ────────────────────────────────────────────

def _discover_eod_files() -> dict:
    """
    Scan data/eod/ for CSV files and identify each one by matching its
    column headers against EOD_FILE_SIGNATURES. Returns a dict mapping
    role -> {path, columns, filename}.

    Raises MissingInputFileError if:
      - No CSVs found in the directory
      - A role has zero matching files
      - Multiple files match the same role (handled because each assigned
        file is removed from consideration for subsequent roles)
    """
    csv_paths = sorted(config.EOD_DATA_DIR.glob("*.csv"))

    if not csv_paths:
        raise MissingInputFileError(
            f"No CSV files found in {config.EOD_DATA_DIR}. "
            "Place at least the required CSV files there and run again."
        )

    # Load column headers for every CSV in the folder.
    candidates = {}
    for path in csv_paths:
        df = pd.read_csv(path, nrows=0)
        candidates[path.name] = {
            "path": path,
            "columns": set(df.columns),
        }

    # Match each role to the best file (highest signature-column overlap).
    assigned = {}
    used_files = set()

    for role, sig in config.EOD_FILE_SIGNATURES.items():
        best_file = None
        best_score = 0

        for fname, info in candidates.items():
            if fname in used_files:
                continue
            score = len(info["columns"] & sig)
            if score > best_score:
                best_score = score
                best_file = fname

        if best_file is None or best_score == 0:
            found_list = ", ".join(candidates)
            raise MissingInputFileError(
                f"Could not find a CSV with the '{role}' signature columns: "
                f"{', '.join(sorted(sig))}.\n"
                f"Found files: [{found_list}]\n"
                f"Ensure a CSV in {config.EOD_DATA_DIR} has the "
                f"expected columns listed above."
            )

        assigned[role] = {
            "path": candidates[best_file]["path"],
            "columns": candidates[best_file]["columns"],
            "filename": best_file,
        }
        used_files.add(best_file)

    # Warn about unmatched files that will be ignored.
    for fname in candidates:
        if fname not in used_files:
            print(f"Warning: '{fname}' did not match any known "
                  f"signature and will be ignored.")

    return assigned


def _validate_eod_headers(role: str, columns: set, filename: str):
    """
    Check that a discovered file has all required columns for its role.
    Raises MissingHeaderError with a clear message if any are missing.
    """
    required = config.EOD_REQUIRED_COLUMNS[role]
    missing = [col for col in required if col not in columns]

    if missing:
        msg_lines = [
            "",
            "=" * 60,
            f"MISSING REQUIRED HEADERS in '{filename}' (matched as '{role}')",
            "=" * 60,
            f"Expected: {', '.join(required)}",
            f"Found:    {', '.join(sorted(columns))}",
            f"Missing:  {', '.join(missing)}",
            "",
            "Fix: make sure the CSV file has the columns listed above",
            "with those exact names, then run the script again.",
            "=" * 60,
            "",
        ]
        raise MissingHeaderError("\n".join(msg_lines))


def discover_eod_file_paths() -> dict:
    """Public: {role: Path} for the EOD input files currently matched in
    data/eod/. Used after a successful run to know what to archive."""
    return {role: info["path"] for role, info in _discover_eod_files().items()}


def load_all() -> dict:
    """
    Discover, validate, and load all 3 EOD input CSVs.
    Returns dict with keys 'conversations', 'kpi_results', 'twilio_events'.
    """
    discovered = _discover_eod_files()

    loaded = {}
    for role, info in discovered.items():
        _validate_eod_headers(role, info["columns"], info["filename"])
        with Spinner(f"Loading {info['filename']}"):
            loaded[role] = pd.read_csv(info["path"])

    return loaded


# ── Mode 2: Contact List ──────────────────────────────────────────

def resolve_contact_list_path(path=None):
    """
    Resolve the contact list input file path.

    If a path is given (via --input): verify it exists and return it.
    If not: scan data/contact_list/ and discover the file by matching
    its column headers against REQUIRED_CONTACT_LIST_HEADERS.

    Returns a Path to the file. Raises MissingInputFileError on failure.
    """
    if path is not None:
        target = Path(path) if not isinstance(path, Path) else path
        if not target.exists():
            message = "\n".join([
                "",
                "=" * 60,
                "MISSING INPUT FILE - cannot generate the contact list.",
                "=" * 60,
                f"Expected file: {target}",
                "",
                "Fix: check that the file exists at the path above,",
                "or omit --input to auto-discover in data/contact_list/.",
                "=" * 60,
                "",
            ])
            raise MissingInputFileError(message)
        return target

    # Auto-discover: scan data/contact_list/ for files with matching headers.
    # Only CSV/Excel are supported — a bare .txt phone list can't carry the
    # customer name the agent's spiel needs, or the application date.
    csv_signature = set(config.REQUIRED_CONTACT_LIST_HEADERS)
    matches = []

    for ext in ("*.csv", "*.xlsx", "*.xls"):
        for fpath in sorted(config.CONTACT_LIST_DATA_DIR.glob(ext)):
            try:
                if fpath.suffix.lower() == ".csv":
                    df = pd.read_csv(fpath, nrows=0)
                else:
                    df = pd.read_excel(fpath, nrows=0)
                if csv_signature.issubset(set(df.columns)):
                    matches.append(fpath)
            except Exception:
                continue

    if len(matches) == 1:
        return matches[0]

    if len(matches) > 1:
        names = ", ".join(f"'{f.name}'" for f in matches)
        message = "\n".join([
            "",
            "=" * 60,
            "AMBIGUOUS - multiple files match the contact list headers.",
            "=" * 60,
            f"Found: {names}",
            "",
            "All these files are recognized as valid contact list inputs",
            f"(CSV/Excel with {', '.join(config.REQUIRED_CONTACT_LIST_HEADERS)} columns).",
            "Use --input <path> to specify which one to use.",
            "=" * 60,
            "",
        ])
        raise MissingInputFileError(message)

    # Zero matches — list everything in the directory for clarity.
    all_entries = sorted(
        f.name for f in config.CONTACT_LIST_DATA_DIR.glob("*")
        if f.is_file() and f.name != ".gitkeep"
    )
    found = ", ".join(all_entries) if all_entries else "(empty directory)"
    message = "\n".join([
        "",
        "=" * 60,
        "MISSING INPUT FILE - cannot generate the contact list.",
        "=" * 60,
        f"Expected headers: {', '.join(config.REQUIRED_CONTACT_LIST_HEADERS)}",
        f"Scanned folder:   {config.CONTACT_LIST_DATA_DIR}",
        f"Files found:      {found}",
        "",
        "Fix: place a CSV or Excel file with the columns listed above in",
        "the folder above, or use --input <path> to point at a specific",
        "file elsewhere.",
        "=" * 60,
        "",
    ])
    raise MissingInputFileError(message)


def load_contact_list(path) -> pd.DataFrame:
    """Load the raw GFiber contact list (customer_phone, user,
    application_date). Supports .csv and .xlsx/.xls — auto-detected from
    the file extension."""
    with Spinner(f"Loading {path.name}"):
        if path.suffix.lower() == ".csv":
            return pd.read_csv(path)
        return pd.read_excel(path, sheet_name=0)


def validate_contact_list_headers(df: pd.DataFrame):
    """
    Checks that the DataFrame contains all required columns.
    Raises MissingHeaderError with a clear message if any are missing.
    """
    required = config.REQUIRED_CONTACT_LIST_HEADERS
    missing = [col for col in required if col not in df.columns]

    if missing:
        message = "\n".join([
            "",
            "=" * 60,
            "MISSING REQUIRED HEADERS - cannot generate the contact list.",
            "=" * 60,
            f"Expected headers: {', '.join(required)}",
            f"Found headers:    {', '.join(df.columns.tolist())}",
            f"Missing header(s): {', '.join(missing)}",
            "",
            "Fix: make sure the input file (CSV or Excel) has the columns",
            "listed above with those exact names, then run the script again.",
            "=" * 60,
            "",
        ])
        raise MissingHeaderError(message)
