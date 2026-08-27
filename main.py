"""
main.py
--------
Entry point. Run this file to generate either:
  - Mode 1 "eod": GFiber Application EOD Report + Call Detail Log workbook
  - Mode 2 "contact-list": GFiber Application Contact List workbook

Usage:
    python main.py --mode eod --agent-id 1595                                                       # EOD mode, today (PHT)
    python main.py --mode eod --agent-id 1595 --start-date 2026-08-25 --end-date 2026-08-29         # EOD mode, a date range

    python main.py --mode contact-list                                  # Contact List mode, stamped with today (PHT)
    python main.py --mode contact-list --as-of-date 2026-08-26          # Contact List mode, stamped with a specific date
    python main.py --mode contact-list --input path/to/other_list.xlsx  # override the input file
"""

import argparse
import datetime
import sys

import pandas as pd

from src import config, preprocessing, call_detail, eod_report, excel_writer, validators, contact_list, data_loader, validation_report, prior_day, archiver
from src.data_loader import MissingInputFileError, MissingHeaderError


def parse_args():
    parser = argparse.ArgumentParser(description="Generate GFiber Abandoned Application reports.")
    parser.add_argument(
        "--mode", choices=["eod", "contact-list"], required=True,
        help="Which report to generate: 'eod' (EOD Report + Call Detail Log) "
             "or 'contact-list' (GFiber Application Contact List from the customer list workbook).",
    )

    # --- Mode 1 (eod) args ---
    parser.add_argument(
        "--start-date", type=str, default=None,
        help="[eod mode] Start of the report period, format YYYY-MM-DD. Must be given together with --end-date. "
             "Omit both to default to today (PHT).",
    )
    parser.add_argument(
        "--end-date", type=str, default=None,
        help="[eod mode] End of the report period, format YYYY-MM-DD (inclusive). Must be given together with --start-date.",
    )
    parser.add_argument(
        "--agent-id", type=int, default=None,
        help="[eod mode] Agent ID to report on. Required when --mode eod.",
    )

    # --- Mode 2 (contact-list) args ---
    parser.add_argument(
        "--as-of-date", type=str, default=None,
        help="[contact-list mode] Report date (YYYY-MM-DD) used to stamp the output folder "
             "and filenames. Defaults to today in PHT.",
    )
    parser.add_argument(
        "--input", type=str, default=None,
        help="[contact-list mode] Path to the contact list workbook. Omit to auto-discover in data/contact_list/.",
    )

    return parser.parse_args()


# ── Mode 1: EOD Report ────────────────────────────────────────────

def run_eod(agent_id: int, start_date=None, end_date=None):
    # 1. Load + clean + merge the 3 source files, filtered to this agent.
    working_table = preprocessing.build_working_table(agent_id=agent_id)

    if working_table.empty:
        print(f"No conversations found for agent_id={agent_id}. Nothing to report.")
        sys.exit(1)

    # 2. Build the Call Detail Log (all calls, not date-filtered yet).
    detail_log = call_detail.build_call_detail_log(working_table)

    # 3. Default to today (PHT) if no range was given.
    if start_date is None:
        start_date = end_date = pd.Timestamp.now(tz=config.TIMEZONE).date()
        print(f"No --start-date/--end-date given, defaulting to today (PHT): {start_date}")

    # 3b. Make sure the resolved period actually has data before reporting on it.
    has_data_in_range = (
        (detail_log["Call Date (PHT)"] >= start_date) & (detail_log["Call Date (PHT)"] <= end_date)
    ).any()
    if not has_data_in_range:
        period_label = str(start_date) if start_date == end_date else f"{start_date} to {end_date}"
        print(f"No calls found for agent_id={agent_id} in the period {period_label}. Nothing to report.")
        sys.exit(1)

    # 4. Build the aggregated EOD summary covering the whole period.
    eod_df = eod_report.build_eod_report(detail_log, start_date, end_date, agent_id)

    # 5. Slice the detail log down to the report period for the sheet.
    range_detail_log = detail_log[
        (detail_log["Call Date (PHT)"] >= start_date) & (detail_log["Call Date (PHT)"] <= end_date)
    ].reset_index(drop=True)

    # 6. Write both sheets to a date- and run-time-stamped subfolder inside output/eod/.
    run_time = datetime.datetime.now()
    eod_dir = config.get_eod_run_output_dir(start_date, end_date, run_time)
    eod_dir.mkdir(parents=True, exist_ok=True)

    if start_date == end_date:
        report_filename = config.OUTPUT_FILENAME_TEMPLATE_SINGLE.format(agent_id=agent_id, start_date=start_date)
        val_filename = config.EOD_VALIDATION_FILENAME_TEMPLATE_SINGLE.format(agent_id=agent_id, start_date=start_date)
    else:
        report_filename = config.OUTPUT_FILENAME_TEMPLATE_RANGE.format(agent_id=agent_id, start_date=start_date, end_date=end_date)
        val_filename = config.EOD_VALIDATION_FILENAME_TEMPLATE_RANGE.format(agent_id=agent_id, start_date=start_date, end_date=end_date)

    # 6b. Look up yesterday's saved report (single-day runs only) to fill
    #     in the "Yesterday" column.
    previous_day_values = {}
    if start_date == end_date:
        previous_date = start_date - datetime.timedelta(days=1)
        previous_day_values = prior_day.load_previous_day_values(agent_id, previous_date)

    report_path = excel_writer.resolve_output_path(eod_dir / report_filename)
    excel_writer.write_eod_report_sheets(eod_df, range_detail_log, report_path, previous_day_values=previous_day_values)
    print(f"EOD report generated: {report_path}")

    # 7. Build and write the validation report alongside the EOD report.
    val_sheets = validation_report.build_validation_report(
        working_table, detail_log, eod_df, start_date, end_date, agent_id,
    )
    val_path = excel_writer.resolve_output_path(eod_dir / val_filename)
    excel_writer.write_validation_report(val_sheets, val_path)
    print(f"Validation report generated: {val_path}")

    # 8. Archive the processed input files so data/eod/ is ready for the next drop.
    matched_paths = data_loader.discover_eod_file_paths()
    archive_dir = config.get_eod_archive_dir(start_date, end_date, run_time)
    archiver.archive_files(list(matched_paths.values()), archive_dir)
    print(f"Input files archived to: {archive_dir}")

    return report_path


# ── Mode 2: Contact List ──────────────────────────────────────────

def run_contact_list(as_of_date=None, input_path=None):
    # 1. Resolve path (auto-discover or explicit --input), then load.
    path = data_loader.resolve_contact_list_path(input_path)
    raw_df = data_loader.load_contact_list(path)

    if raw_df.empty:
        print("Contact list is empty. Nothing to report.")
        sys.exit(1)

    # 2. Validate required headers.
    data_loader.validate_contact_list_headers(raw_df)

    # 3. Default the report date to today in PHT if not given. It only
    #    stamps the output folder and filenames.
    if as_of_date is None:
        as_of_date = pd.Timestamp.now(tz=config.TIMEZONE).date()
        print(f"No --as-of-date given, defaulting to today (PHT): {as_of_date}")

    # 4. Validate and categorize every record.
    categories = contact_list.categorize_records(raw_df)

    # 5. Build summary statistics.
    total = len(raw_df)
    valid_count = len(categories["valid"])
    invalid_count = len(categories["invalid"])

    def pct(n):
        return round(n / total * 100, 1) if total else 0.0

    summary_df = pd.DataFrame({
        "Metric": [
            "Total Records",
            "Valid",
            "Invalid",
        ],
        "Count": [total, valid_count, invalid_count],
        "% of Total": [
            100.0,
            pct(valid_count),
            pct(invalid_count),
        ],
    })

    # 6. Write outputs inside a date- and run-time-stamped subfolder.
    run_time = datetime.datetime.now()
    output_dir = config.get_contact_list_run_output_dir(as_of_date, run_time)
    output_dir.mkdir(parents=True, exist_ok=True)

    contact_path = None

    # All valid records for CSV output, in input file order.
    all_records = categories["valid"].copy()
    all_records["ref_id"] = config.CONTACT_LIST_REF_ID

    if not all_records.empty:
        # 6a. Write Contact List CSV.
        filename = config.CONTACT_LIST_OUTPUT_FILENAME_TEMPLATE.format(date=as_of_date)
        contact_path = excel_writer.resolve_output_path(output_dir / filename)
        excel_writer.write_contact_list_csv(all_records, contact_path)
        print(f"Contact list generated: {contact_path}")
    else:
        print("No valid records found. Contact list not generated.")

    # 6b. Write Validation Report (2-sheet workbook).
    validation_filename = config.VALIDATION_OUTPUT_FILENAME_TEMPLATE.format(date=as_of_date)
    validation_path = excel_writer.resolve_output_path(output_dir / validation_filename)

    sheets = {
        "summary": summary_df,
        "invalid": categories["invalid"],
    }
    excel_writer.write_validation_report(sheets, validation_path)
    print(f"Validation report generated: {validation_path}")

    # 7. Archive the input file (only if it was auto-discovered, not an
    #    explicit --input override) so data/contact_list/ is ready for next time.
    if input_path is None:
        archive_dir = config.get_contact_list_archive_dir(as_of_date, run_time)
        archiver.archive_files([path], archive_dir)
        print(f"Input file archived to: {archive_dir}")

    return contact_path


if __name__ == "__main__":
    args = parse_args()

    if args.mode == "contact-list":
        try:
            as_of = validators.parse_single_date(args.as_of_date) if args.as_of_date else None
        except validators.InvalidDateRangeError as e:
            print(f"Invalid --as-of-date: {e}")
            sys.exit(1)

        try:
            run_contact_list(as_of_date=as_of, input_path=args.input)
        except (MissingInputFileError, MissingHeaderError) as e:
            print(e)
            sys.exit(1)

    else:  # args.mode == "eod"
        try:
            agent_id = validators.require_agent_id(args.agent_id)
        except validators.MissingAgentIdError as e:
            print(e)
            sys.exit(1)

        try:
            parsed_start, parsed_end = validators.parse_date_range(args.start_date, args.end_date)
        except validators.InvalidDateRangeError as e:
            print(f"Invalid date range: {e}")
            sys.exit(1)

        try:
            run_eod(agent_id=agent_id, start_date=parsed_start, end_date=parsed_end)
        except (MissingInputFileError, MissingHeaderError) as e:
            print(e)
            sys.exit(1)
