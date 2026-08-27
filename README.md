# GFiber Abandoned Application Voicebot — Report Automation (EOD Report + Contact List + Validation Report)

Reporting automation for Globe **UC-2: GFiber Abandoned Application** — the
outbound campaign that follows up with customers who started a Globe At Home
G Fiber 1499 application but never finished it.

Generates one of two reports:

- **EOD Report** — a 2-sheet Excel workbook (**EOD Report** + **Call Detail
  Log**) from 3 raw source CSVs (`conversations`, `kpi_results`,
  `twilio_webhook_events`), filtered to a single agent and a calling-day
  range. Also writes a companion **5-sheet Validation Report** workbook
  alongside it (join/data-quality diagnostics for that same run).

- **Contact List** — produces one CSV and one Excel workbook from the
  customer list file provided by Globe: a **Contact List CSV** (valid
  records, oldest abandoned application first) and a **Validation Report**
  (2-sheet data quality report with invalid records).

---

## 1. Step-by-step: setting this up in VS Code (from scratch)

1. **Install prerequisites** (skip if already installed):
   - [Python 3.10+](https://www.python.org/downloads/)
   - [VS Code](https://code.visualstudio.com/)
   - In VS Code, install the **Python extension** (Microsoft) from the Extensions panel (`Ctrl+Shift+X` / `Cmd+Shift+X`, search "Python").

2. **Open the project folder in VS Code**
   `File → Open Folder...` → select the `gfiber-application-automation` folder.

3. **Open a terminal inside VS Code**
   `Terminal → New Terminal` (or `` Ctrl+` ``).

4. **Create a virtual environment** (see "Do we need venv?" below for why):
```bash
   python -m venv venv
```

5. **Activate it**
   - Windows (PowerShell): `venv\Scripts\Activate.ps1`
   - Mac/Linux: `source venv/bin/activate`

   VS Code may also prompt "Select Interpreter" — pick the one inside
   `./venv`. Do this via `Ctrl+Shift+P` → "Python: Select Interpreter" if it
   doesn't prompt automatically.

6. **Install dependencies**
```bash
   pip install -r requirements.txt
```

7. **Add your input files** into the appropriate subdirectories under `data/`.
   Files are auto-discovered by matching their column headers, so they can be
   named anything — the names below are just examples:

   **EOD mode** (3 CSVs in `data/eod/`, one matching each required column set):
   - `data/eod/conversations.csv` (needs `conversation_id`, `agent_id`, `start_timestamp`, `end_timestamp`, `call_logs`, `contact_number`)
   - `data/eod/kpi_results.csv` (needs `voiceConversationId`, `voiceAgentId`, `outputJson`)
   - `data/eod/twilio_webhook_events.csv` (needs `conversation_id`, `event`)

   **Contact List mode** (one file in `data/contact_list/`):
   - CSV or Excel, needs `customer_phone`, `user`, and `application_date` columns:
     `data/contact_list/gfiber_abandoned_applications.xlsx`

8. **Run it**
```bash
   python main.py --mode eod --agent-id 1595                             # EOD mode, today (PHT)
   python main.py --mode contact-list                                    # Contact List mode, as-of today (PHT)
   python main.py --mode contact-list --as-of-date 2026-08-26            # Contact List, specific date
   python main.py --mode contact-list --input path/to/other.xlsx         # Contact List, override input file
```
   The generated files land in `output/eod/{date-or-range}/{HH-MM-SS}/` (EOD mode) or
   `output/contact_list/{date}/{HH-MM-SS}/` (Contact List mode) — a date-stamped
   subfolder for the report period, with a further run-time-stamped subfolder so
   reruns for the same period never mix their files together.
   On a successful run, the input file(s) that were used are also moved out of
   `data/eod/` or `data/contact_list/` into `archive/` (see below), so the
   input folder is empty and ready for the next drop.

---

## 2. Tools used

| Tool | Purpose |
|---|---|
| Python 3.10+ | Core language |
| `pandas` | CSV loading, cleaning, joining, aggregation |
| `openpyxl` | Writing formatted Excel workbooks (EOD Report: 2-sheet, EOD Validation Report: 5-sheet, Contact List Validation Report: 2-sheet) |
| `pandas.to_csv` | Writing the Contact List CSV |
| VS Code + Python extension | Editor / debugger |
| `venv` (built into Python, no separate install) | Isolated dependency environment |

## Do we need venv?

**Yes, recommended.** It keeps `pandas`/`openpyxl` versions for this project
separate from anything else on your machine, so upgrading a package for a
different project later can't silently break this script (or vice versa).
It's a couple of extra terminal commands (steps 4–5 above) for meaningfully
safer long-term maintenance — worth it even for a small script like this.

---

## 3. Project structure (modular by design)

```
gfiber-application-automation/
├── data/
│   ├── contact_list/          → input: any CSV/Excel with customer_phone + user + application_date columns (Contact List mode)
│   └── eod/                   → input: 3 CSVs, auto-matched by column headers (EOD mode)
├── output/
│   ├── contact_list/
│   │   └── {date}/            → generated: GFiber_Application_Contact_List_{date}.csv
│   │                            + GFiber_Application_Validation_Report_{date}.xlsx
│   └── eod/
│       └── {date-or-range}/   → generated: GFiber_Application_EOD_Report_{agent_id}_{date}.xlsx
│                                + GFiber_Application_EOD_Validation_{agent_id}_{date}.xlsx
├── archive/                    → processed input files, moved here after a successful run
│   ├── contact_list/
│   │   └── {date}_{HHMMSS}/   → the contact list file used for that run
│   └── eod/
│       └── {date-or-range}_{HHMMSS}/   → the 3 source CSVs used for that run
├── src/
│   ├── config.py              → all settings in one place (paths, timezone, filename templates, required headers, ref_id)
│   ├── validators.py          → validates date CLI args (--start-date/--end-date, --as-of-date) and --agent-id
│   ├── data_loader.py         → auto-discovers input files by column headers; validates required headers are present
│   ├── archiver.py            → moves processed input files into archive/ after a successful run
│   ├── progress.py            → spinning "loading..." animation in the terminal
│   ├── preprocessing.py       → cleans data, parses JSON, filters to one agent, joins 3 sources (EOD mode)
│   ├── call_detail.py         → builds the "Call Detail Log" sheet (one row per call, one column per KPI field)
│   ├── eod_report.py          → builds the "EOD Report" sheet (aggregated funnel summary)
│   ├── prior_day.py           → best-effort reads back yesterday's saved EOD workbook to fill in the Yesterday/Δ columns (single-day runs only)
│   ├── contact_list.py        → builds the Contact List + validates/categorizes records (valid, invalid); normalizes phone numbers to +63XXXXXXXXXX
│   ├── validation_report.py   → builds the EOD mode's 5-sheet Validation Report (Join Summary, Field Completeness, Calculation Audit, Data Quality Issues, Duplicate Contacts)
│   └── excel_writer.py        → writes DataFrames to formatted .xlsx (EOD Report, EOD Validation Report, Contact List Validation Report) and the Contact List to .csv
├── main.py                    → entry point; dispatches to run_eod() or run_contact_list() based on --mode
├── requirements.txt           → pandas, openpyxl
└── README.md                  → this file
```

Each module has exactly one job, so you can swap any piece (e.g. point
`data_loader.py` at a database instead of CSVs later) without touching the
others.

---

## 4. Usage

Choose a mode with `--mode` (required — either `eod` or `contact-list`):

### Mode 1: EOD Report (`--mode eod`)

```bash
python main.py --mode eod --agent-id 1595                                                   # today (PHT)
python main.py --mode eod --agent-id 1595 --start-date 2026-08-25 --end-date 2026-08-29     # a date range
python main.py --mode eod --agent-id 1595 --start-date 2026-08-29 --end-date 2026-08-29     # a single day (range of 1)
```

- `--start-date` / `--end-date`: inclusive range in `YYYY-MM-DD`. Must be
  given together (or both omitted — defaults to today in PHT). Start cannot
  be after end.
- `--agent-id`: required for `--mode eod`. No default — the run fails with
  an error if omitted. The GFiber agent is `1595` in Dev.

**Output naming:** single-day → `GFiber_Application_EOD_Report_{agent_id}_{date}.xlsx`;
multi-day → `GFiber_Application_EOD_Report_{agent_id}_{start}_to_{end}.xlsx`.
Lands in `output/eod/{date-or-range}/{HH-MM-SS}/`, where `{HH-MM-SS}` is the
run's clock time — each run of the tool gets its own subfolder, even for the
same date-or-range.

**Report structure:** the workbook has 2 sheets. **EOD Report** is a
dashboard-style summary — one row per metric for the whole period (not one
row per day) — with a Today/Yesterday/Δ comparison table, grouped into
sections: call volume, **FUNNEL** (identity → consent → postpaid),
**OUTCOMES** (intent, endorsement, leads, final dispositions),
**NON-COMPLETION & COMPETITOR**, **QUALITY**, then the manually-filled
FINOPS / ISSUES & CHANGES / TOMORROW'S PLAN blocks.

For single-day runs, **Yesterday** is best-effort filled in by reading back
the previous day's already-generated EOD workbook (see `src/prior_day.py`) —
it degrades to blank if no prior report exists for that date, or it can't be
read. **Δ** becomes a live Excel formula for any row where Yesterday got
populated. Multi-day (range) runs don't get a prior-day lookup, since
"yesterday" isn't well-defined for a range — those columns stay blank.
Everything above the FINOPS section participates in the Today/Yesterday/Δ
comparison; FINOPS onward is filled in by hand and shows a `[+/-N]`
placeholder instead. The "Day X of 14" subtitle is left as a placeholder,
since the pipeline doesn't track which day of the campaign a run
corresponds to. Cells that need a human to fill them in day-to-day (FinOps,
issues/changes) are highlighted yellow.

**Call Detail Log** lists every call in the period across 30 columns — one
per field the agent emits in `outputJson`, plus the Twilio status, duration,
and a `Call Date (PHT)` column so you can still see which day each row
belongs to.

**Counting rule:** call-volume counts (Dialed, Connected, No Answer, Busy,
Failed) come from every row in the period. Every KPI-derived count
(identity, consent, postpaid, intent, endorsement, competitor, quality) is
counted from **connected calls only**, so the funnel and the rates built on
it stay honest — a no-answer call can't have confirmed an identity or stated
an intent.

**Also generated:** a companion `GFiber_Application_EOD_Validation_{agent_id}_{date}.xlsx`
workbook is written alongside the EOD Report on every run, in the same
output folder — a 5-sheet diagnostics report (Join Summary, Field
Completeness, Calculation Audit, Data Quality Issues, Duplicate Contacts)
covering that same period's source data. The **Calculation Audit** sheet
recomputes every metric independently and marks each `PASS` or `MISMATCH`
against what the EOD Report actually wrote.

**Input archiving:** once both files above have been written successfully,
the 3 source CSVs used for that run are moved out of `data/eod/` into
`archive/eod/{date-or-range}_{HHMMSS}/`, so `data/eod/` is empty and ready
for the next drop. If the run fails or exits early (no data found, missing
headers, etc.), the input files are left in place untouched so you can fix
and re-run without digging through the archive.

### Mode 2: Contact List (`--mode contact-list`)

```bash
python main.py --mode contact-list                               # as-of today (PHT)
python main.py --mode contact-list --as-of-date 2026-08-26       # specific reference date
python main.py --mode contact-list --input path/to/other.csv     # override input file (CSV or Excel)
```

- `--as-of-date`: reference date in `YYYY-MM-DD` used to compute
  `days_since_applied`. Defaults to today in PHT.
- `--input`: override the contact list file path. If omitted, the file is
  auto-discovered in `data/contact_list/` by matching column headers
  (`customer_phone` + `user` + `application_date`). Supports `.csv`,
  `.xlsx`, and `.xls`.

**Why `user` is required:** the agent's opening message and every
re-verification / closing spiel interpolate `{user}` — the customer's name.
A blank name would put a broken sentence on the call, so records missing one
are rejected rather than dialled.

**Header validation:** before any processing, the script checks that the input
file contains all three required columns. If any is missing, processing stops
immediately with a clear error message.

**Phone normalization:** valid phone numbers in all outputs are formatted as
`+63XXXXXXXXXX` (no spaces, consistent prefix) regardless of the input format.
Three input formats are accepted, all normalizing to the same `+63XXXXXXXXXX`:
- `+63`/`63` + 10 digits (12 digits total), e.g. `+63 998 766 5432` → `+639987665432`
- `09` + 9 digits (11 digits total), e.g. `09987665432` → `+639987665432`
- `9` + 9 digits (10 digits total), e.g. `9987665432` → `+639987665432`

**Output:** two files in `output/contact_list/{date}/{HH-MM-SS}/` (a date-stamped
subfolder, with a further run-time-stamped subfolder per run):

- `GFiber_Application_Contact_List_{date}.csv` — every record that passes
  validation, with `customer_phone` normalized to `+63XXXXXXXXXX`, sorted by
  `days_since_applied` **descending** (oldest abandoned application first —
  the longest-neglected applications get dialled first). Includes a `ref_id`
  column (constant value from `config.CONTACT_LIST_REF_ID`, currently
  `GOCUC20`). `days_since_applied` is a plain calendar-day difference (e.g.
  as-of date Aug 26, `application_date` Aug 10 → `16`).
- `GFiber_Application_Validation_Report_{date}.xlsx` — data validation +
  categorization (2 sheets: Summary, Invalid Data)

**Input archiving:** once both files above have been written successfully,
the contact list file used for that run is moved out of
`data/contact_list/` into `archive/contact_list/{date}_{HHMMSS}/`, so
`data/contact_list/` is empty and ready for the next drop. This only
applies when the file was auto-discovered — if you used `--input` to point
at a file elsewhere, it's left where it is. If the run fails or exits early,
the input file is left in place untouched.

**Validation rules** (applied to every record):

| # | Check | Fails when... |
|---|---|---|
| 1 | Missing phone | `customer_phone` is blank |
| 2 | Invalid PH code | Number doesn't start with `+63`, `63`, `09`, or `9` |
| 3 | Invalid length | Wrong digit count for the matched prefix (12 for `+63`/`63`, 11 for `09`, 10 for `9`), after stripping non-digits |
| 4 | Missing customer name | `user` is blank |
| 5 | Missing date | `application_date` is blank |
| 6 | Invalid date | `application_date` cannot be parsed |
| 7 | Duplicate phone | Same `customer_phone` appears more than once |

Checks 1–3 are a phone chain (stops at the first failure — e.g. a missing
phone won't also report an invalid code); check 4 (name) and checks 5–6 (date
chain) run independently of the phone chain and of each other. Check 7
(duplicate) always runs globally, in addition to whatever the other checks
found. Multiple reasons on the same row are joined with `"; "`.

**Validation Report — 2 sheets:**

| Sheet | Content |
|---|---|
| Summary | Record counts per category with percentages |
| Invalid Data | Rows that failed validation with concatenated `reason` column |

---

## 5. Known data caveats (read before trusting the numbers)

- **Status is sourced exclusively from the Twilio call-progress journey**
  (`twilio_webhook_events.csv`'s `event` column), never from
  `conversations.status`. If a conversation's `conversation_id` has no
  matching row in `twilio_webhook_events.csv`, **Status is left blank** —
  it is not guessed or backfilled from anything else. That's expected, not
  a bug — it'll start populating automatically, no code changes needed,
  once Twilio coverage exists for the agent being reported on.
- **Call duration is sourced from `call_logs.metrics.total_duration_ms`,
  not `end_timestamp`.** `end_timestamp` frequently ends up identical to
  `start_timestamp` (zero duration) or earlier than it (negative duration),
  and doesn't correlate with the real call length recorded in `call_logs`.
  It's epoch-millisecond UTC like `start_timestamp` and converts to a
  technically valid date, but the *value itself* isn't a trustworthy
  call-end moment — so it's excluded from duration math.
  (`start_timestamp` is reliable, and is what drives the Call Date /
  `--start-date`/`--end-date` filtering.)
- **Contact Number is partially recoverable, not fully.** Some rows arrive
  as Excel scientific notation (e.g. `"6.39178E+11"`), which only keeps
  5-6 significant digits — the real trailing digits were lost *upstream*,
  before this script ever sees the data, and can't be reconstructed from
  this file alone. The working table tracks a contact-number reliability
  label per row, surfaced in the Validation Report's **Data Quality
  Issues** sheet:
  - `Complete` — the number arrived uncorrupted and was normalized to
    `63XXXXXXXXXX`.
  - `Complete (recovered from Twilio)` — the number was corrupted in
    `conversations.csv`, but `twilio_webhook_events.csv` had a matching
    call with the true number in its `To` field, so that was used
    instead.
  - `TRUNCATED - only first N digits are real, rest lost upstream` — no
    Twilio match existed to recover it, so the (zero-padded, incomplete)
    number is shown as-is rather than presented as if it were complete.
    Don't use these for actual outreach without going back to Globe's
    source system.
- **Same-day duplicate dials are collapsed.** The Call Detail Log holds one
  row per (Contact Number, Call Date) — a `Connected` attempt always wins
  outright, otherwise the latest attempt by Call Time is kept. Every row
  that took part in a collision is listed in the Validation Report's
  **Duplicate Contacts** sheet, showing which one was kept.
- **Contact List mode validates every record.** Invalid records (bad phone
  format, missing name, unparseable date, duplicates) are written to the
  **Invalid Data** sheet of the Validation Report with a specific reason.
  There is no age cutoff — every record that passes validation goes on the
  list, however old the application. Note that a **future-dated**
  `application_date` therefore still passes validation; it produces a
  negative `days_since_applied` and sorts to the bottom of the list, but it
  is not rejected. Treat a negative value as an upstream export error worth
  raising with Globe.
- **`call_logs` schema varies by agent.** Some agents store it as
  `{"metrics": {"total_duration_ms": ...}}`. Others store it as a list of
  turn-by-turn bot/user events instead — a different shape entirely.
  Duration extraction handles this defensively (returns blank rather than
  crashing) but doesn't currently parse that alternate schema for duration;
  that's a known gap if/when this pipeline is pointed at those agents.
- **The agent's "Provider Availed" KPI is bound to a field it never emits.**
  In the agent config, that KPI's `outputFieldName` is `provider_availed`,
  but no such key exists among the 24 `conversation_metrics` — the real
  field carrying the provider is `competitor_name`. This report's
  **Provider Availed - PLDT** row reads `competitor_name`, so the number
  here is correct; the agent's own KPI dashboard is the thing reading
  nothing. Worth raising with Globe to get the binding fixed.
- **LLM Inference Cost, P0/P1 issue counts, and several other EOD Report
  fields** are left blank — no source data currently supports them.
  This is separate from `N/A`, which only appears in the Call
  Detail Log's "Application Completed" column when KPI data is
  missing for that call.
