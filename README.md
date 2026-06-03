# NILE Pipeline Issue Tracker

> **Automated monthly pipeline issue reporting for the NILE Platform Production Support team.**
> Reads Jira service desk exports, categorizes pipeline issues using keyword intelligence, and generates a professional multi-month Excel report — ready for manager, director, and VP review.

---

## Table of Contents

- [Project Goal](#project-goal)
- [Background and Purpose](#background-and-purpose)
- [How It Works](#how-it-works)
- [Project Structure](#project-structure)
- [Issue Categories](#issue-categories)
- [Tech Stack](#tech-stack)
- [First Time Setup](#first-time-setup)
- [Monthly Workflow](#monthly-workflow)
- [Output Report Structure](#output-report-structure)
- [How the Excel Report Is Self-Authoring](#how-the-excel-report-is-self-authoring)
- [How to Add or Update Keywords](#how-to-add-or-update-keywords)
- [How to Add a New Category](#how-to-add-a-new-category)
- [Troubleshooting](#troubleshooting)
- [Project Build Journey](#project-build-journey)
- [Author](#author)

---

## Project Goal

The NILE Platform Production Support team handles hundreds of Jira service desk tickets every month. At the end of each month, the team manually counted, categorized, and reported these tickets to management — a process that took hours and was prone to inconsistency.

This project **fully automates that process**. Drop a Jira export file into a folder, run one command, and a professional Excel report is generated in seconds — with category counts, month-over-month trends, charts, and an executive summary ready to share with managers, directors, and VPs.

---

## Background and Purpose

### Who uses this

The **NILE Platform Production Support team** at Cotiviti Nepal Pvt. Ltd. The team supports the NILE data platform which processes healthcare claims, eligibility, and pharmacy data for multiple payer clients across the United States.

### What problem it solves

Every month:
- Hundreds of Jira tickets are created by teams across the NILE platform (Hadoop team, DB team, Platform Engineering, Production Support)
- Not all tickets are genuine pipeline issues — many are change requests, copy operations, access requests, and admin tasks
- The team needed a consistent, accurate count of **genuine pipeline issues** broken down by category
- This count was previously done manually in Excel — time-consuming and inconsistent month to month

### What this tool does

1. Reads the monthly Jira export (HTML, CSV, or XLSX format)
2. Reads all ticket fields — Summary, Description, Comments, Close Notes
3. Uses a **weighted keyword scoring engine** to categorize each ticket into the correct issue category
4. Uses **Close Notes and Comments with higher weight** (3x and 2x) than Summary/Description (1x) — because the resolution tells us more than the symptom
5. Applies **exclusion patterns** to filter out non-issue tickets (admin tasks, duplicates, copy requests, OLA monitoring, transition documents)
6. Generates a growing **multi-month Excel report** — each month adds a new tab without touching previous months
7. All counts in the report are **live Excel formulas** — editing a category via dropdown instantly updates all tabs

---

## How It Works

### The pipeline

```
Input file (HTML / CSV / XLSX)
        ↓
file_finder.py     → auto-detects newest file in input/
        ↓
csv_reader.py      → parses all ticket fields (Summary, Description,
                      Comment, Close Notes, Categories)
        ↓
categorizer.py     → weighted keyword scoring
                      Close Notes  = 3x weight  (resolution = most reliable)
                      Comment      = 2x weight  (what was done)
                      Summary      = 1x weight  (symptom reported)
                      Description  = 1x weight  (symptom detail)
                   → exclusion filter runs first
                      (duplicates, admin tasks, OLA monitoring → Other)
        ↓
excel_writer.py    → writes 4 sheet types per month:
                      Cover Sheet, Summary, Raw Tickets, Needs Review
                   → All Months Summary updated with all months
        ↓
output/Nile_Pipeline_Issue_Report.xlsx   ← grows every month
output/Nile_<Month>_archive.csv          ← flat CSV backup
output/run_<Month>_<timestamp>.log       ← audit log
```

### Weighted scoring explained

The categorizer counts how many keywords from each category match across all ticket fields. It picks the category with the highest score. The weighting ensures that if a ticket says "data drop" in the Description (symptom) but "update the schema" in Close Notes (resolution) — it correctly gets categorized as Schema Issue, not Metric Drop.

### Exclusion filter

Before scoring, the categorizer checks Close Notes and Comments for patterns that indicate the ticket is not a genuine pipeline issue:
- Duplicate tickets ("duplicate ticket created", "another ticket already opened")
- OLA monitoring tasks ("buffer time has been added, I will monitor")
- Admin operations (renaming spoolers, removing pipelines, transition documents)
- Client-side issues ("I realized this was on the client side")
- BMC review/peer review tasks ("listed all the BMC optimization pipelines")

These are sent directly to **Other** without scoring.

---

## Project Structure

```
nile-pipeline-issue-count-report-v3/
│
├── input/                              ← Drop your Jira export here each month
│   └── .gitkeep
│
├── output/                             ← Reports generated here automatically
│   └── .gitkeep
│
├── nile_pipeline_tracker/              ← Core source code
│   ├── __init__.py
│   ├── config.py                       ← Categories, keywords, exclusion patterns
│   ├── categorizer.py                  ← Weighted keyword scoring + exclusion filter
│   ├── csv_reader.py                   ← HTML / CSV / XLSX parser
│   ├── excel_writer.py                 ← All Excel sheet builders
│   ├── file_finder.py                  ← Auto-detects newest file in input/
│   └── main.py                         ← Orchestrates the full pipeline
│
├── run.py                              ← THE file you run every month
├── requirements.txt                    ← Python dependencies
├── README.md                           ← This file
└── .gitignore                          ← Keeps data files off GitHub
```

> **Note:** `input/` and `output/` contents are never pushed to GitHub. Only code goes to GitHub. Each machine maintains its own local data.

---

## Issue Categories

The tracker categorizes tickets into these 11 pipeline issue categories:

| # | Category | What it covers |
|---|----------|----------------|
| 1 | **Configuration Issue** | GWF flag issues, parameter changes, mapping fixes, pipeline suspension, SFTP password changes, skip ingestion date |
| 2 | **File Issue** | Corrupt/invalid/empty files, sequence number problems, wrong filename, replacement file needed, file format changes |
| 3 | **Spooler Issue** | Spooler failures, files not being spooled, ingestion on hold, backfill operations, NAS spooler failures |
| 4 | **Data Issue** | Data mismatch, null values, duplicate data, incorrect data, data validation failures, date span issues |
| 5 | **File Missing** | Files not received from client, history files missing, pipeline on hold waiting for files |
| 6 | **Unmatched Raw Data Count** | Count mismatch between raw and model, lineage report discrepancies, record count differences |
| 7 | **Pipeline Schedule** | Pipeline failures, getting failed, ConvertExcelToCSV failures, permission denied, broadcast timeout, BMC pipeline failures |
| 8 | **DB Issue** | Database errors, disk space issues, transaction log issues, HDFS/Hive issues, DB writer failures |
| 9 | **Metric Drop** | Low volume, metric percentage drop, record counts low, no claim volume, override for metric drop |
| 10 | **Schema Issue** | Schema mismatch, column mismatch, layout update, blank field values, new column added, field not being sent |
| 11 | **Issue Platform** | SFTP/FTP issues, platform down, server issues, Axway issues, connectivity issues |
| — | **Other** | Anything not matching the above — visible in Needs Review tab for manual correction |

---

## Tech Stack

| Component | Technology |
|-----------|-----------|
| Language | Python 3.12 |
| Excel generation | openpyxl 3.1+ |
| HTML parsing | BeautifulSoup4 + lxml |
| CSV/XLSX reading | Python csv (built-in) + openpyxl |
| Keyword matching | Python re (regex, built-in) |
| Logging | Python logging (built-in) |
| Runtime | Local machine — no server, no cloud, no database |

No paid software, no licenses, no internet connection required to run.

---

## First Time Setup

### Prerequisites

- Python 3.12 or later — download from [python.org](https://www.python.org/downloads/)
- Git — download from [git-scm.com](https://git-scm.com/)
- VS Code (recommended) or any terminal

### Step 1 — Clone the repository

```bash
git clone https://github.com/Birajan12345/Nile-pipeline-issue-count-report-v3.git
cd Nile-pipeline-issue-count-report-v3
```

### Step 2 — Install Python dependencies

```bash
pip install -r requirements.txt
```

This installs: `openpyxl`, `beautifulsoup4`, `lxml`

### Step 3 — Verify the setup

```bash
python -c "from nile_pipeline_tracker.main import main; print('Setup OK')"
```

If you see `Setup OK` — you are ready to run.

### Step 4 — Create input and output folders (if not present)

```bash
mkdir -p input output
```

---

## Monthly Workflow

This is the only thing you do every month — 3 steps.

### Step 1 — Export from Jira

Export the month's service desk tickets from Jira as:
- **HTML** — File → Save Page As in your browser (recommended — captures all fields)
- **CSV** — Jira export feature
- **XLSX** — Excel export from Jira

Name the file something meaningful:
```
Service Desk_April.html
Service Desk_May.html
```

### Step 2 — Drop the file into input/

```
input/
└── Service Desk_May.html    ← just drop it here
```

Old files from previous months can stay — the tracker always picks the **most recently modified** file automatically.

### Step 3 — Run the tracker

**Option A — VS Code (recommended)**
Press `F5` or `Ctrl+Shift+B`

**Option B — Terminal / Git Bash**
```bash
python run.py
```

**Option C — Specific file**
```bash
python run.py "input/Service Desk_May.html"
```

**Option D — Windows (if `python` not found)**
```bash
py run.py "input/Service Desk_May.html"
```

### That's it

The report is in `output/Nile_Pipeline_Issue_Report.xlsx` — updated with the new month's data.

---

## Output Report Structure

The report is a single Excel file that grows every month:

```
Nile_Pipeline_Issue_Report.xlsx
│
├── Cover Sheet              ← Always shows the latest month — opens first
├── All Months Summary       ← Side-by-side comparison of all months
├── Summary — Apr 2026       ← April category counts + chart + vs last month
├── Raw — Apr 2026           ← All April tickets with editable category dropdown
├── Needs Review — Apr 2026  ← April uncategorized tickets for manual review
├── Summary — May 2026       ← May category counts + chart + vs last month
├── Raw — May 2026           ← All May tickets with editable category dropdown
├── Needs Review — May 2026  ← May uncategorized tickets for manual review
└── _categories              ← Hidden — powers the dropdown in Raw tabs
```

### Cover Sheet

The first sheet that opens. Designed for VP and director review:
- Branded header: **NILE Platform — Production Support**
- Report month and auto-generated label
- 4 KPI cards: Total Issues, Top 2 categories, Needs Review %
- Health status bar (amber warning if Other rate > 40%, green if healthy)
- Category snapshot table with ▲/▼ trend arrows vs previous month
- Mini bar chart per category
- Footer note

### All Months Summary

Side-by-side view of all months processed so far:
- Rows = issue categories
- Columns = each month (Apr count, Apr %, May count, May %)
- All counts are **live COUNTIF formulas** — update when you edit Raw tabs
- Bar chart showing all months side by side

> **Note:** After opening the file, press `Ctrl+Alt+F9` to force Excel to recalculate all formulas. Then save. This is only needed once per machine.

### Monthly Summary tabs (one per month)

For each month:
- Category name, count, % share, vs last month (▲/▼), color tag
- Bar chart (one bar per category, colored by category)
- Pie chart (top categories distribution)
- All counts are live COUNTIF formulas referencing the Raw tab

### Raw Tickets tabs (one per month)

- Every ticket with all fields: Key, Summary, Description, Comment, Close Notes, Status, Created, Category, Review Required
- **Category column has a dropdown** — click any cell to change the category
- Changing a category here instantly updates the Summary tab and All Months Summary

### Needs Review tabs (one per month)

- Lists all tickets categorized as "Other"
- Review these and use the dropdown in the Raw tab to recategorize obvious ones
- Instruction note at bottom: "Go to Raw tab → find ticket → use Category dropdown"

---

## How the Excel Report Is Self-Authoring

The report is designed so that **manual corrections flow through automatically**:

```
You change a category in Raw — May 2026 tab via dropdown
        ↓
Summary — May 2026 count updates instantly (COUNTIF formula)
        ↓
All Months Summary May column updates instantly (COUNTIF formula)
        ↓
Cover Sheet KPIs update instantly (linked to Summary)
        ↓
Charts update automatically
```

No need to re-run the tracker after manual corrections. The Excel formulas handle everything.

---

## How to Add or Update Keywords

All keywords are in `nile_pipeline_tracker/config.py` in the `CATEGORIES` list.

To add a new keyword to an existing category:

```python
{
    "name": "Spooler Issue",
    "color": "FF0000",
    "keywords": [
        "spooler", "spoolers", "not able to spool",
        # Add your new keyword here:
        "your new keyword phrase",
    ],
},
```

**Rules for good keywords:**
- Use **2+ word phrases** — single words are too broad and cause wrong matches
- Use the **exact phrasing your team uses** in ticket summaries and close notes
- Never add a keyword that could match tickets from a different category

After changing keywords — re-run the tracker with the same input file to regenerate the report with updated categorization.

---

## How to Add a New Category

1. Open `nile_pipeline_tracker/config.py`
2. Add a new entry to the `CATEGORIES` list:

```python
{
    "name": "My New Category",      # exact name shown in report
    "color": "AABBCC",              # hex color for charts (without #)
    "keywords": [
        "keyword one",
        "keyword two phrase",
        "another specific phrase",
    ],
},
```

3. Save the file
4. Re-run the tracker

The new category will automatically appear in all Summary tabs, All Months Summary, Cover Sheet, and the Raw tab dropdown.

---

## Troubleshooting

### `Python was not found`
Try:
```bash
py run.py "input/Service Desk_May.html"
```
Or open the project in VS Code and press `F5`.

### `No supported file found in 'input/'`
The input folder is empty. Drop your Jira export HTML/CSV/XLSX file into the `input/` folder first.

### `Sheet 'Raw — May 2026' already exists — skipping`
This month was already processed. If you want to reprocess it, delete `output/Nile_Pipeline_Issue_Report.xlsx` and re-run both months.

### All Months Summary shows blank counts
Open the file → click **Enable Content** → press **Ctrl+Alt+F9** → save. This forces Excel to recalculate the COUNTIF formulas.

### Security warning on open
Click **Enable Content**. Then go to **File → Options → Trust Center → Trust Center Settings → Trusted Locations** and add your `output/` folder. The warning will never appear again.

### `ModuleNotFoundError: No module named 'bs4'`
Run:
```bash
pip install -r requirements.txt
```

---

## Project Build Journey

This project was built in 7 phases, each adding a layer of capability:

| Phase | What was built |
|-------|---------------|
| **Phase 1** | Core categorization engine — 11 categories, keyword scoring, config-driven |
| **Phase 2** | Folder-aware auto-run — input/ folder detection, XLSX support, run logging |
| **Phase 3** | Executive Excel report — Cover Sheet, KPI cards, health status, trend arrows |
| **Phase 4** | Accuracy tuning — weighted field scoring (Close Notes 3x, Comment 2x), exclusion filter for admin/duplicate/client-side tickets, keyword expansion from real ticket analysis |
| **Phase 5** | Monthly Summary visual improvements — horizontal bar chart, data labels, color-coded categories, % Share column |
| **Phase 6** | Multi-month architecture — one growing Excel file, per-month Raw/Summary/Needs Review tabs, All Months Summary with live COUNTIF formulas, vs Last Month trend column |
| **Phase 7** | GitHub setup and documentation |

### Key design decisions

**Why weighted scoring?**
The first version used first-match logic — it stopped at the first keyword match. This caused tickets to land in the wrong category when their summary mentioned one issue but their resolution was about something different. Weighted scoring counts all keyword matches across all fields and picks the highest scorer. Close Notes (the resolution) carries 3x weight because what was actually fixed is more reliable than how the problem was initially described.

**Why an exclusion filter?**
Not all 200+ tickets per month are genuine pipeline issues. Admin tasks, file copy operations, transition documents, duplicate tickets, and client-side non-issues would inflate the counts. The exclusion filter checks Close Notes and Comments for resolution patterns that indicate non-issues — and sends them directly to Other before any keyword scoring runs.

**Why one growing Excel file instead of one per month?**
The team needs to see trends over time. Having one file per month means manually comparing them. A single growing file with all months lets the All Months Summary tab show the full trend picture automatically. Each month's data is isolated in its own tabs — nothing overwrites previous months.

**Why COUNTIF formulas instead of hardcoded values?**
The categorizer is not perfect — some tickets need manual correction after review. With hardcoded values, a correction in the Raw tab would not update the Summary. With COUNTIF formulas, any change to a category via the dropdown in Raw immediately reflects across Summary, All Months Summary, and Cover Sheet without re-running the tracker.

---

## Author

**Birajan** — NILE Platform Production Support Team  
Cotiviti Nepal Pvt. Ltd.

Built to eliminate manual monthly reporting and give the production support team a professional, accurate, and automated pipeline issue tracking system.

---

*This project is for internal use by the NILE Platform Production Support team. Input files (Jira exports) and output reports (Excel files) are never committed to GitHub — they remain local to each machine.*
