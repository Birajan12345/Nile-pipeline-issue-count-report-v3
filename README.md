# NILE Pipeline Issue Tracker

This project takes a Jira CSV export and generates an Excel report with:
- raw ticket details,
- a summary by category,
- charts,
- and a "Needs Review" sheet for uncategorized issues.

## Project structure

- `nile_pipeline_csv_tracker.py`
  - The main launcher script.
  - It imports and runs `nile_pipeline_tracker.main.main()`.
  - Use this file to execute the tracker directly.

- `nile_pipeline_tracker/`
  - The refactored package containing the tracker logic.

  - `__init__.py`
    - Exposes `main` from the package.

  - `main.py`
    - Orchestrates the workflow.
    - Reads the CSV file.
    - Parses and categorizes rows.
    - Generates the Excel workbook.
    - Writes the report to the output folder.

  - `config.py`
    - Holds environment and CSV configuration.
    - Defines column aliases used to map Jira export headers.
    - Defines issue categories, colors, and keyword matching rules.

  - `csv_reader.py`
    - Implements CSV loading and normalization.
    - Detects column names using `COLUMN_ALIASES`.
    - Parses Jira dates into `YYYY-MM-DD`.
    - Builds the row dictionary used by the report.
    - Extracts the report month label from ticket dates.

  - `categorizer.py`
    - Contains the text-based category matcher.
    - Searches summaries and descriptions for keywords.
    - Returns the first matching category or `Other / Uncategorized`.

  - `excel_writer.py`
    - Builds the Excel workbook using `openpyxl`.
    - Creates three sheets:
      - `Raw Tickets`
      - `Monthly Summary`
      - `Needs Review`
    - Styles headers, formats columns, and adds charts.

## How it works

1. `nile_pipeline_csv_tracker.py` launches the package.
2. `main.py` loads the CSV from `CSV_PATH` or the first command-line argument.
3. `csv_reader.py` reads the file, detects column names, and converts dates.
4. `categorizer.py` assigns a category based on the summary and description.
5. `main.py` passes parsed rows to `excel_writer.py`.
6. `excel_writer.py` writes the workbook and saves the file.

## Requirements

- Python 3.12
- `openpyxl`
- `beautifulsoup4`
- `lxml`

Install dependencies with:

```powershell
& "c:/Users/birajan.raya/Projects/Project Issue count/.venv/Scripts/python.exe" -m pip install openpyxl beautifulsoup4 lxml
```

## Run the tracker

From the project folder:

```powershell
& "c:/Users/birajan.raya/Projects/Project Issue count/.venv/Scripts/python.exe" nile_pipeline_csv_tracker.py sample_jira_export.csv
```

Or parse a Microsoft Edge exported Jira HTML file:

```powershell
& "c:/Users/birajan.raya/Projects/Project Issue count/.venv/Scripts/python.exe" nile_pipeline_csv_tracker.py "Service Desk.html"
```

Or use `CSV_PATH` in `nile_pipeline_tracker/config.py`:

```python
CSV_PATH = "jira_export.csv"
```

## Customization

- Update keywords or categories in `nile_pipeline_tracker/config.py`.
- Change output folder with `OUTPUT_DIR` in `nile_pipeline_tracker/config.py`.
- Use `REPORT_MONTH_LABEL` to force a custom report title.
# Nile-pipeline-issue-count-report-v3
