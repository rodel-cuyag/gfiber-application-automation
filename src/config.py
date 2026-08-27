"""
config.py
---------
Single source of truth for file paths and settings.
"""

from pathlib import Path

# ── Project folders ──────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent.parent      # project root
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"

# Each mode gets its own input/output subfolder so the two features
# never collide or get confused about which file is which.
EOD_DATA_DIR = DATA_DIR / "eod"
EOD_OUTPUT_DIR = OUTPUT_DIR / "eod"

CONTACT_LIST_DATA_DIR = DATA_DIR / "contact_list"
CONTACT_LIST_OUTPUT_DIR = OUTPUT_DIR / "contact_list"

# ── Mode 1: EOD Report — input CSVs ──────────────────────────────
# Header signatures for auto-discovery — data_loader scans data/eod/
# for CSVs and identifies each file by these unique column sets.
# Files can be named anything as long as the required columns are present.
EOD_FILE_SIGNATURES = {
    "conversations": {"agent_id", "call_logs"},
    "kpi_results":   {"voiceConversationId", "outputJson"},
    "twilio_events": {"call_sid", "event"},
}

# Full column requirements validated after signature-based discovery.
EOD_REQUIRED_COLUMNS = {
    "conversations": [
        "conversation_id", "agent_id", "start_timestamp",
        "end_timestamp", "call_logs", "contact_number",
    ],
    "kpi_results": [
        "voiceConversationId", "voiceAgentId", "outputJson",
    ],
    "twilio_events": [
        "conversation_id", "event",
    ],
}

# ── Mode 2: Contact List — input file ────────────────────────────
# data_loader auto-discovers the file in data/contact_list/ by scanning
# for CSV/Excel files that have the required columns below.
# Use --input on the CLI to override with an explicit path.

# ── Output file naming ────────────────────────────────────────────
# Filled in with the report date(s) at runtime (see main.py).
# Single-day EOD runs (start == end) use the plain template; multi-day
# ranges use the range template so the filename itself shows the span.
OUTPUT_FILENAME_TEMPLATE_SINGLE = "GFiber_Application_EOD_Report_{agent_id}_{start_date}.xlsx"
OUTPUT_FILENAME_TEMPLATE_RANGE = "GFiber_Application_EOD_Report_{agent_id}_{start_date}_to_{end_date}.xlsx"

EOD_VALIDATION_FILENAME_TEMPLATE_SINGLE = "GFiber_Application_EOD_Validation_{agent_id}_{start_date}.xlsx"
EOD_VALIDATION_FILENAME_TEMPLATE_RANGE = "GFiber_Application_EOD_Validation_{agent_id}_{start_date}_to_{end_date}.xlsx"

CONTACT_LIST_OUTPUT_FILENAME_TEMPLATE = "GFiber_Application_Contact_List_{date}.csv"
VALIDATION_OUTPUT_FILENAME_TEMPLATE = "GFiber_Application_Validation_Report_{date}.xlsx"

# The GFiber agent's opening message interpolates {user}, so the customer
# name is required alongside the number — a blank name would produce a
# broken spiel on the call.
REQUIRED_CONTACT_LIST_HEADERS = ["customer_phone", "user"]

# Constant ref_id value stamped on every row of the Contact List CSV output.
CONTACT_LIST_REF_ID = "GOCUC20"

# ── EOD output folderization ──────────────────────────────────────
# Each date range gets its own subfolder under output/eod/ to keep
# reports organised as they accumulate.
def get_eod_output_dir(start_date, end_date):
    if start_date == end_date:
        folder_name = str(start_date)
    else:
        folder_name = f"{start_date}_to_{end_date}"
    return EOD_OUTPUT_DIR / folder_name


# Each run within a date folder gets its own time-stamped subfolder, so
# reruns for the same date never mix their files together.
def get_eod_run_output_dir(start_date, end_date, run_time):
    return get_eod_output_dir(start_date, end_date) / f"{run_time:%H-%M-%S}"


def get_contact_list_run_output_dir(as_of_date, run_time):
    return CONTACT_LIST_OUTPUT_DIR / str(as_of_date) / f"{run_time:%H-%M-%S}"


# ── Input archiving ────────────────────────────────────────────────
# After a successful run, processed input files are moved out of data/eod/
# and data/contact_list/ into dated folders here, so the input folder
# empties out and is ready for the next drop. Folder names combine the
# report's business date (matching the output folder above) with the
# run's clock time, so re-runs never collide or overwrite each other.
ARCHIVE_DIR = BASE_DIR / "archive"
EOD_ARCHIVE_DIR = ARCHIVE_DIR / "eod"
CONTACT_LIST_ARCHIVE_DIR = ARCHIVE_DIR / "contact_list"


def get_eod_archive_dir(start_date, end_date, run_time):
    folder_name = str(start_date) if start_date == end_date else f"{start_date}_to_{end_date}"
    return EOD_ARCHIVE_DIR / f"{folder_name}_{run_time:%H%M%S}"


def get_contact_list_archive_dir(as_of_date, run_time):
    return CONTACT_LIST_ARCHIVE_DIR / f"{as_of_date}_{run_time:%H%M%S}"


# ── Timezone ──────────────────────────────────────────────────────
# Source timestamps are epoch millis (UTC). Plan requires PHT (UTC+8).
TIMEZONE = "Asia/Manila"
