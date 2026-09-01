"""
excel_writer.py
-----------------
Writes report DataFrames into formatted .xlsx workbooks using openpyxl.

Note: values here are pre-computed in Python (not live Excel formulas).
These are operational snapshot reports generated fresh each run from
source data, rather than financial models meant to be edited live in
Excel — so static, correct values are the right choice here.
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
import pandas as pd

HEADER_FILL = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF")
BODY_FONT = Font(name="Arial")

SECTION_HEADER_FILL = PatternFill("solid", start_color="D6E4F0", end_color="D6E4F0")
SECTION_HEADER_FONT = Font(name="Arial", bold=True, size=11)

YELLOW_FILL = PatternFill("solid", start_color="FFFF00", end_color="FFFF00")

# ── "EOD Report" sheet (single dashboard-style summary sheet) ─────

DASH_FONT_NAME = "Work Sans"
DASH_NAVY = "1A2C52"
DASH_LIGHT_BLUE = "E8F0F9"
DASH_GRAY = "555555"
DASH_BLUE_TEXT = "0066CC"
DASH_PURPLE_TEXT = "7B5EA7"
DASH_GREEN_FILL = "00B050"
DASH_RED_FILL = "C00000"

_THIN = Side(style="thin")
_FULL_BORDER = Border(top=_THIN, bottom=_THIN, left=_THIN, right=_THIN)


def _dash_row(label, source=None, highlight=False, yellow=False, no_delta=False, special=None, percent_delta=False):
    return {
        "label": label, "source": source, "highlight": highlight,
        "yellow": yellow, "no_delta": no_delta, "special": special,
        "percent_delta": percent_delta,
    }


# One entry per sheet row: display label, the matching eod_df metric label
# to pull the "Today" value from (written as a literal value, not a
# formula), and any special styling. "Failed" and "Total Completed Calls"
# are included even though the reference dashboard doesn't show them, for
# consistency with the metrics computed elsewhere (see plan doc).
DASHBOARD_ROWS = [
    _dash_row("Calls Dialled - Target", "Calls Dialed - Target", yellow=True),
    _dash_row("Calls Dialled - Actual", "Calls Dialed - Actual"),
    _dash_row("Calls Connected", "Calls Connected", highlight=True),
    _dash_row("No Answer", "No Answer"),
    _dash_row("Busy", "Busy"),
    _dash_row("Failed", "Failed"),
    _dash_row("System Errors", special="system_errors"),
    _dash_row("Contacted", "Contacted"),
    _dash_row("Total Completed Calls", "Total Completed Calls"),
    _dash_row("Total Call Duration (minutes)", "Total Call Duration (minutes)"),
    _dash_row("Avg. Call Duration - Connected (seconds)", "Avg. Call Duration - Connected (seconds)"),
    _dash_row("Connection Rate (Connected / Dialled)", "Connection Rate (Connected / Dialed)", highlight=True, percent_delta=True),
    _dash_row("__SECTION__", special="FUNNEL"),
    _dash_row("Identity Confirmed", "Identity Confirmed", highlight=True),
    _dash_row("Wrong Customer", "Wrong Customer"),
    _dash_row("Consented to Continue and Recording", "Consented to Continue and Recording", highlight=True),
    _dash_row("Declined Recording", "Declined Recording"),
    _dash_row("Postpaid Verified", "Postpaid Verified"),
    _dash_row("Postpaid Customers % (of Connected)", "Postpaid Customers % (of Connected)", percent_delta=True),
    _dash_row("Non-Postpaid Verified", "Non-Postpaid Verified"),
    _dash_row("Non-Postpaid Customers % (of Connected)", "Non-Postpaid Customers % (of Connected)", percent_delta=True),
    _dash_row("__SECTION__", special="OUTCOMES"),
    _dash_row("Wishes to Proceed", "Wishes to Proceed", highlight=True),
    _dash_row("Wishes to Proceed - Postpaid", "Wishes to Proceed - Postpaid"),
    _dash_row("Wishes to Proceed - Non-Postpaid", "Wishes to Proceed - Non-Postpaid"),
    _dash_row("No Longer Interested", "No Longer Interested"),
    _dash_row("No Longer Interested - Postpaid", "No Longer Interested - Postpaid"),
    _dash_row("No Longer Interested - Non-Postpaid", "No Longer Interested - Non-Postpaid"),
    _dash_row("Application Already Completed", "Application Already Completed"),
    _dash_row("Application Already Completed - Postpaid", "Application Already Completed - Postpaid"),
    _dash_row("Application Already Completed - Non-Postpaid", "Application Already Completed - Non-Postpaid"),
    _dash_row("Conversion Rate (Proceed / Connected)", "Conversion Rate (Proceed / Connected)", highlight=True, percent_delta=True),
    _dash_row("Endorsed for Work Order", "Endorsed for Work Order", highlight=True),
    _dash_row("Endorsed for Work Order % (of Connected)", "Endorsed for Work Order % (of Connected)", percent_delta=True),
    _dash_row("Lead for Outbound Handling", "Lead for Outbound Handling"),
    _dash_row("Lead for Outbound Handling % (of Connected)", "Lead for Outbound Handling % (of Connected)", percent_delta=True),
    _dash_row("Email Remarketing Tagged", "Email Remarketing Tagged"),
    _dash_row("Email Remarketing % (of Connected)", "Email Remarketing % (of Connected)", percent_delta=True),
    _dash_row("Postpaid Conversion", "Postpaid Conversion"),
    _dash_row("Non-Postpaid Conversion", "Non-Postpaid Conversion"),
    _dash_row("Not Available / No Consent", "Not Available / No Consent"),
    _dash_row("Retries Queued for Tomorrow", "Retries Queued for Tomorrow"),
    _dash_row("__SECTION__", special="NON-COMPLETION & COMPETITOR"),
    _dash_row("Non-Completion - Price", "Non-Completion - Price"),
    _dash_row("Non-Completion - Competitor", "Non-Completion - Competitor"),
    _dash_row("Competitor Identified", "Competitor Identified"),
    _dash_row("__SECTION__", special="QUALITY"),
    _dash_row("Repeat Requested (quality)", "Repeat Requested (quality)"),
    _dash_row("Identity Re-asked (defect)", "Identity Re-asked (defect)"),
    _dash_row("Opt-Out Requested", "Opt-Out Requested"),
    _dash_row("__SECTION__", special="FINOPS"),
    _dash_row("LLM Inference Cost", "LLM Inference Cost (USD)", yellow=True),
    _dash_row("Total Daily Spend", "Total Daily Spend (USD)", yellow=True, highlight=True),
    _dash_row("__SECTION__", special="ISSUES & CHANGES"),
    _dash_row("Open P0 issues", "Open P0 Issues", yellow=True),
    _dash_row("Open P1 issues", "Open P1 Issues", yellow=True),
    _dash_row("Changes deployed today", "Changes Deployed Today", yellow=True, no_delta=True),
    _dash_row("Changes pending approval for tomorrow", "Changes Pending Approval for Tomorrow", yellow=True, no_delta=True),
    _dash_row("__SECTION__", special="TOMORROW'S PLAN"),
    _dash_row("Target call volume", "Target Call Volume", no_delta=True, yellow=True),
    _dash_row("Expected list from Globe (ETA)", "Expected List from Globe (ETA)", no_delta=True, yellow=True),
    _dash_row("Calling window", "Calling Window", no_delta=True),
    _dash_row("Phase gate status", "Phase Gate Status", no_delta=True, special="phase_gate", yellow=True),
]

# Rows that get comma-thousands, no-decimal number formatting ("#,##0") —
# the count metrics only. Averages, percentage rates (which are literal
# "84.3%" strings, not numeric), FINOPS $ amounts, and issue counts are
# intentionally excluded.
COUNT_FORMAT_SOURCES = {
    "Calls Dialed - Target", "Calls Dialed - Actual", "Calls Connected",
    "No Answer", "Busy", "Failed", "Contacted", "Total Completed Calls",
    "Identity Confirmed", "Wrong Customer",
    "Consented to Continue and Recording", "Declined Recording",
    "Postpaid Verified", "Non-Postpaid Verified",
    "Wishes to Proceed",
    "Wishes to Proceed - Postpaid", "Wishes to Proceed - Non-Postpaid",
    "No Longer Interested",
    "No Longer Interested - Postpaid", "No Longer Interested - Non-Postpaid",
    "Application Already Completed",
    "Application Already Completed - Postpaid",
    "Application Already Completed - Non-Postpaid",
    "Endorsed for Work Order", "Lead for Outbound Handling",
    "Email Remarketing Tagged",
    "Postpaid Conversion", "Non-Postpaid Conversion",
    "Not Available / No Consent", "Retries Queued for Tomorrow",
    "Non-Completion - Price", "Non-Completion - Competitor",
    "Competitor Identified",
    "Repeat Requested (quality)", "Identity Re-asked (defect)",
    "Opt-Out Requested",
}

# Call Detail Log column widths, keyed by column NAME rather than letter so
# the mapping survives any reordering of the log's 30 columns. Unlisted
# columns keep the auto-computed width from _write_dataframe.
CALL_DETAIL_LOG_COLUMN_WIDTHS = {
    "Conversation ID": 40,
    "Contact Number": 18,
    "Customer Sentiment & Feedback": 45,
    "Customer Questions & Concerns": 45,
    "Question Topics": 45,
}


def resolve_output_path(path: Path) -> Path:
    """
    If *path* already exists, append *(1)*, *(2)*, etc. before the
    extension so the existing file is never silently overwritten.
    """
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    n = 1
    while True:
        candidate = parent / f"{stem} ({n}){suffix}"
        if not candidate.exists():
            return candidate
        n += 1


def _write_dataframe(ws, df: pd.DataFrame):
    """Writes a DataFrame to a worksheet with a styled header row and auto-width columns."""
    ws.append(list(df.columns))
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    for row in df.itertuples(index=False):
        ws.append(list(row))

    for i, col in enumerate(df.columns, start=1):
        if not df.empty:
            col_len = df[col].astype(str).str.len().max()
            col_len = 0 if pd.isna(col_len) else col_len
        else:
            col_len = 0
        max_len = max(col_len, len(str(col)))
        ws.column_dimensions[get_column_letter(i)].width = min(max_len + 4, 45)

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = BODY_FONT

    ws.freeze_panes = "A2"


def _write_eod_summary_sheet(ws, eod_df: pd.DataFrame, previous_day_values: dict = None):
    """
    Writes the "EOD Report" sheet: big navy header, Today/Yesterday/Delta
    comparison table. "Today" values are literal values computed in Python
    (from eod_df), not formulas — there's no separate source sheet to pull
    from. "System Errors" is the one live formula, referencing the Target/
    Actual rows on this same sheet.

    Yesterday (Column C) is filled in for the core dashboard rows (5-18)
    from *previous_day_values* (a {label: value} lookup from the previous
    day's saved report — see prior_day.py) when available, else left
    blank. Delta (Column D) becomes a live Excel formula for those same
    rows once C is populated; rows from the FINOPS section onward, and
    the "Day X of N" campaign counter, stay blank/bracket placeholders —
    those aren't derived from prior-day report data.
    """
    value_of = dict(eod_df.itertuples(index=False))
    previous_day_values = previous_day_values or {}
    # The "core block" — everything with a real prior-day value to compare
    # against — runs up to FINOPS, not up to the first section header. The
    # funnel/outcome/quality sections above FINOPS are all data-derived and
    # must keep their Yesterday/Delta columns; FINOPS onward is filled in by
    # hand, so those rows get the "[+/-N]" placeholder instead.
    first_section_idx = next(
        i for i, s in enumerate(DASHBOARD_ROWS) if s["special"] == "FINOPS"
    )

    ws.merge_cells("A1:D1")
    title = ws["A1"]
    title.value = "End-of-Day Report"
    title.font = Font(name=DASH_FONT_NAME, size=36, color=DASH_NAVY)
    title.alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 42

    subtitle = ws["A2"]
    # 14 is the fixed known campaign length; the current day number isn't
    # tracked anywhere yet, so it stays a fill-in-by-hand placeholder.
    subtitle.value = f"Date: {value_of['Report Period']}    Day: [N] of 14"
    subtitle.font = Font(name=DASH_FONT_NAME, size=18, bold=True, color=DASH_GRAY)

    navy_fill = PatternFill("solid", start_color=DASH_NAVY, end_color=DASH_NAVY)
    header_font = Font(name=DASH_FONT_NAME, size=18, bold=True, color="FFFFFF")
    for col, text in enumerate(["Metric", "Today", "Yesterday", "Δ"], start=1):
        cell = ws.cell(row=3, column=col, value=text)
        cell.font = header_font
        cell.fill = navy_fill
    ws.cell(row=3, column=1).border = Border(top=_THIN, left=_THIN)
    ws.cell(row=3, column=2).border = Border(top=_THIN)
    ws.cell(row=3, column=3).border = Border(top=_THIN)
    ws.cell(row=3, column=4).border = Border(top=_THIN, right=_THIN)

    subheader_font = Font(name=DASH_FONT_NAME, size=18, color="F5F5F5")
    for col, text in enumerate(["", "T0", "T-1", "vs T-1"], start=1):
        cell = ws.cell(row=4, column=col, value=text)
        cell.font = subheader_font
        cell.fill = navy_fill
    ws.cell(row=4, column=1).border = Border(left=_THIN)
    ws.cell(row=4, column=4).border = Border(right=_THIN)

    def _fill(color):
        return PatternFill("solid", start_color=color, end_color=color)

    def _available(label):
        return previous_day_values.get(label) not in (None, "")

    row_of = {}
    system_errors_row = None
    body_row = 5
    for idx, spec in enumerate(DASHBOARD_ROWS):
        if spec["label"] == "__SECTION__":
            ws.merge_cells(start_row=body_row, start_column=2, end_row=body_row, end_column=4)
            a = ws.cell(row=body_row, column=1)
            a.fill = navy_fill
            a.border = Border(left=_THIN)
            b = ws.cell(row=body_row, column=2, value=spec["special"])
            b.font = header_font
            for col in (2, 3, 4):
                ws.cell(row=body_row, column=col).fill = navy_fill
            ws.cell(row=body_row, column=4).border = Border(right=_THIN)
            body_row += 1
            continue

        if spec["source"]:
            row_of[spec["source"]] = body_row

        highlight = spec["highlight"]
        band_color = DASH_LIGHT_BLUE if highlight else "FFFFFF"
        label_font = Font(name=DASH_FONT_NAME, size=18, color=DASH_GRAY)
        value_font = Font(
            name=DASH_FONT_NAME, size=18, bold=highlight,
            color=DASH_BLUE_TEXT if highlight else DASH_NAVY,
        )
        delta_font = Font(name=DASH_FONT_NAME, size=18, color=DASH_PURPLE_TEXT)

        a = ws.cell(row=body_row, column=1, value=spec["label"])
        a.font = label_font
        a.fill = _fill(band_color)
        a.border = _FULL_BORDER

        if spec["special"] == "system_errors":
            value = f"=MAX(0,B{row_of['Calls Dialed - Target']}-B{row_of['Calls Dialed - Actual']})"
            system_errors_row = body_row
        elif spec["source"] == "Retries Queued for Tomorrow":
            base_retries = value_of[spec["source"]]
            value = f"=MAX(0,{base_retries}+B{system_errors_row})"
        else:
            value = value_of[spec["source"]]
        is_count_row = spec["source"] in COUNT_FORMAT_SOURCES or spec["special"] == "system_errors"

        b = ws.cell(row=body_row, column=2, value=value)
        b.font = value_font
        b.fill = YELLOW_FILL if spec["yellow"] else _fill(band_color)
        b.alignment = Alignment(horizontal="left")
        b.border = _FULL_BORDER
        if is_count_row:
            b.number_format = "#,##0"

        in_core_block = idx < first_section_idx

        yesterday_value = None
        if spec["special"] == "system_errors":
            if _available("Calls Dialled - Actual"):
                yesterday_value = (
                    f"=MAX(0,C{row_of['Calls Dialed - Target']}-C{row_of['Calls Dialed - Actual']})"
                )
        elif spec["source"] == "Retries Queued for Tomorrow":
            if all(_available(lbl) for lbl in ("No Answer", "Busy", "Failed")):
                yesterday_value = (
                    f"=MAX(0,C{row_of['No Answer']}+C{row_of['Busy']}+C{row_of['Failed']}+C{system_errors_row})"
                )
        elif in_core_block:
            candidate = previous_day_values.get(spec["label"])
            if candidate not in (None, ""):
                yesterday_value = candidate

        c = ws.cell(row=body_row, column=3, value=yesterday_value)
        c.font = delta_font
        c.fill = _fill(band_color)
        c.alignment = Alignment(horizontal="left")
        c.border = _FULL_BORDER
        if is_count_row:
            c.number_format = "#,##0"

        if spec["no_delta"]:
            delta_value = "—"
        elif not in_core_block:
            delta_value = "[+/-N]"
        elif spec["percent_delta"]:
            delta_value = (
                f'=IFERROR(TEXT(VALUE(SUBSTITUTE(B{body_row},"%",""))'
                f'-VALUE(SUBSTITUTE(C{body_row},"%","")),"+0.0;-0.0")&"%","")'
            )
        else:
            delta_value = f'=IF(AND(ISNUMBER(B{body_row}),ISNUMBER(C{body_row})),B{body_row}-C{body_row},"")'

        d = ws.cell(row=body_row, column=4, value=delta_value)
        d.font = delta_font
        d.fill = _fill(band_color)
        d.alignment = Alignment(horizontal="left")
        d.border = _FULL_BORDER
        if is_count_row:
            d.number_format = "#,##0"

        if spec["special"] == "phase_gate":
            ws.conditional_formatting.add(
                f"B{body_row}",
                FormulaRule(
                    formula=[f'ISNUMBER(SEARCH("GO",B{body_row}))'],
                    fill=_fill(DASH_GREEN_FILL),
                    font=Font(name=DASH_FONT_NAME, bold=True, color="FFFFFF"),
                ),
            )
            ws.conditional_formatting.add(
                f"B{body_row}",
                FormulaRule(
                    formula=[f'ISNUMBER(SEARCH("Hold",B{body_row}))'],
                    fill=_fill(DASH_RED_FILL),
                    font=Font(name=DASH_FONT_NAME, bold=True, color="FFFFFF"),
                ),
            )

        body_row += 1

    ws.column_dimensions["A"].width = 72
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 18.5
    ws.column_dimensions["D"].width = 12


def write_eod_report_sheets(eod_df: pd.DataFrame, call_detail_df: pd.DataFrame, output_path, previous_day_values: dict = None):
    """
    Creates the EOD-mode 2-sheet workbook: 'EOD Report' (dashboard-style
    summary sheet) and 'Call Detail Log'.
    """
    wb = Workbook()

    eod_sheet = wb.active
    eod_sheet.title = "EOD Report"
    _write_eod_summary_sheet(eod_sheet, eod_df, previous_day_values)

    detail_sheet = wb.create_sheet("Call Detail Log")
    columns = list(call_detail_df.columns)
    _write_dataframe(detail_sheet, call_detail_df)
    _apply_date_format(detail_sheet, columns, ["Call Date (PHT)"])
    for col_name, width in CALL_DETAIL_LOG_COLUMN_WIDTHS.items():
        if col_name not in columns:
            continue
        letter = get_column_letter(columns.index(col_name) + 1)
        detail_sheet.column_dimensions[letter].width = width

    wb.save(output_path)
    return output_path


def write_contact_list_sheet(df: pd.DataFrame, output_path, sheet_name: str, date_columns=None):
    """
    Creates a single-sheet workbook for the valid Contact List.
    date_columns: optional list of column names to format as plain
    dates (YYYY-MM-DD) instead of openpyxl's default datetime display.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    _write_dataframe(ws, df)
    _apply_date_format(ws, list(df.columns), date_columns)
    wb.save(output_path)
    return output_path


_SHEET_TITLES = {
    "summary": "Summary",
    "invalid": "Invalid Data",
    # EOD validation sheet titles
    "join_summary": "Join Summary",
    "field_completeness": "Field Completeness",
    "calculation_audit": "Calculation Audit",
    "data_quality_issues": "Data Quality Issues",
    "duplicate_contacts": "Duplicate Contacts",
}


def write_validation_report(sheets: dict, output_path, date_columns=None):
    """
    Creates a multi-sheet validation report workbook.
    Each key in *sheets* is a sheet identifier; the value is a DataFrame.
    The first key determines the first (active) sheet; subsequent keys become
    additional sheets in iteration order.
    """
    keys = list(sheets.keys())
    if not keys:
        raise ValueError("At least one sheet is required")

    wb = Workbook()
    first_key = keys[0]
    wb.active.title = _SHEET_TITLES.get(first_key, first_key.replace("_", " ").title())
    _write_dataframe(wb.active, sheets[first_key])
    if date_columns:
        _apply_date_format(wb.active, list(sheets[first_key].columns), date_columns)

    for key in keys[1:]:
        title = _SHEET_TITLES.get(key, key.replace("_", " ").title())
        ws = wb.create_sheet(title)
        _write_dataframe(ws, sheets[key])
        if date_columns:
            _apply_date_format(ws, list(sheets[key].columns), date_columns)

    wb.save(output_path)
    return output_path


# ── CSV writers (Contact List) ────────────────────────────────────


def write_contact_list_csv(df: pd.DataFrame, output_path):
    """Write the Contact List (valid records) to a CSV file."""
    df.to_csv(output_path, index=False)
    return output_path


# ── Excel helpers ─────────────────────────────────────────────────


def _apply_date_format(ws, column_names, date_columns):
    """Apply YYYY-MM-DD formatting to cells in date_columns (skip header row)."""
    if not date_columns:
        return
    for col_name in date_columns:
        if col_name not in column_names:
            continue
        col_letter = get_column_letter(column_names.index(col_name) + 1)
        for cell in ws[col_letter][1:]:
            cell.number_format = "YYYY-MM-DD"