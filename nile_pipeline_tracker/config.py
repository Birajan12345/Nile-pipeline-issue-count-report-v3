import os

# ── Folder paths ──────────────────────────────────────────
INPUT_DIR  = os.getenv("INPUT_DIR",  "input")
OUTPUT_DIR = os.getenv("OUTPUT_DIR", "output")
SUPPORTED_EXTENSIONS = [".html", ".htm", ".csv", ".xlsx", ".xls"]

# Leave blank → auto-detected from ticket dates each run
REPORT_MONTH_LABEL = os.getenv("REPORT_MONTH_LABEL", "")

COLUMN_ALIASES = {
    "key": ["Issue key", "Key", "Issue Key", "key"],
    "issue_id": ["Issue id", "Issue ID", "ID", "id"],
    "parent_id": ["Parent id", "Parent ID", "Parent", "Parent key"],
    "summary": ["Summary", "summary", "Title"],
    "description": ["Description", "description", "Issue Description"],
    "categories": ["Categories", "categories", "Category", "category", "Catogories", "GSR Category", "GSR Sub Category", "New Categories", "New Catogories", "New Category"],
    "new_categories": ["New Categories", "New Category", "New Catogories", "GSR Sub Category", "GSR Category"],
    "assign": ["Assign", "Assigned To", "Assignee", "assign"],
    "status": ["Status", "status"],
    "created": ["Created", "created", "Date Created", "Created Date"],
    "updated": ["Updated", "updated", "Last Updated", "Updated Date"],
    "assignment_group": ["Assignment Group", "assignment group", "Internal Assignment Group", "Custom field (Assignment Group)"],
    "assignee": ["Assignee", "assignee"],
    "reporter": ["Reporter", "reporter"],
    "resolution": ["Resolution", "resolution"],
    "comments": ["Comments", "comments", "Comment", "customfield_Comments"],
    "close_notes": ["Close Notes", "close notes", "Close Note"],
    "labels": ["Labels", "labels", "Label"],
    "issue_type": ["Issue Type", "Issue type", "issuetype", "Type"],
    "priority": ["Priority", "priority"],
    "inward_cloners": ["Inward issue link (Cloners)"],
    "outward_cloners": ["Outward issue link (Cloners)"],
    "inward_relates": ["Inward issue link (Relates)"],
}

HTML_FIELD_LABELS = {
    "status": ["Status"],
    "issue_type": ["Type", "Issue Type", "GSR Issue Type"],
    "priority": ["Priority"],
    "reporter": ["Reporter"],
    "assignee": ["Assignee"],
    "resolution": ["Resolution"],
    "assignment_group": ["Assignment Group", "Internal Assignment Group"],
    "close_notes": ["Close Notes"],
    "categories": ["Category", "Categories", "GSR Category", "GSR Sub Category"],
    "new_categories": ["New Categories", "GSR Sub Category"],
    "comments": ["Comments"],
    "parent_id": ["Parent", "Parent id", "Parent ID"],
}

CATEGORIES = [
    {
        "name": "Configuration Issue",
        "color": "5B9BD5",
        "keywords": [
            "configuration", "config", "parameter", "setup", "mapping",
            "env issue", "property issue", "config error",
            "configuration issue", "config change", "pipeline config",
            "misconfigured", "config update", "parameter change",
            "pipeline parameter", "wrong config",
        ],
    },
    {
        "name": "File Issue",
        "color": "4472C4",
        "keywords": [
            "corrupt file", "invalid file", "encrypted file",
            "damaged file", "empty file", "bad file", "file issue",
            "file not received", "file not found", "file delay",
            "file arrival", "file drop", "source file", "file rejected",
            "file ingestion", "not ingested", "file failure",
            "files failing", "file failing",
            "path does not exist", "file path does not exist",
        ],
    },
    {
        "name": "Spooler issue",
        "color": "FF0000",
        "keywords": [
            "spooler", "spoolers", "unscheduled", "stuck",
            "trigger issue", "pickup failure", "spool delay",
            "spool error", "spool failed", "spool queue", "spooling",
            "spoolers failed", "not able to spool",
        ],
    },
    {
        "name": "Data Issue",
        "color": "ED7D31",
        "keywords": [
            "data mismatch", "duplicate data", "null values",
            "incorrect data", "missing records", "data issue",
            "data quality", "wrong data", "data discrepancy",
            "data error", "invalid data", "corrupt data", "bad data",
            "data not matching", "data validation", "data integrity",
            "data drop",
        ],
    },
    {
        "name": "File Missing",
        "color": "FFC000",
        "keywords": [
            "file missing", "files missing", "missing file",
            "no file received", "replacement file pending",
            "missing data file", "not posted", "files not posted",
            "not received",
        ],
    },
    {
        "name": "Unmatched raw data count",
        "color": "00B050",
        "keywords": [
            "count mismatch", "unmatched count", "raw count mismatch",
            "record mismatch", "unmatched raw data count",
            "record count mismatch", "row count mismatch",
            "raw data count", "unmatched record",
        ],
    },
    {
        "name": "Pipeline Schedule",
        "color": "FF8C00",
        "keywords": [
            "pipeline failed", "pipelines failed", "workflow failed",
            "cron issue", "delayed workflow", "dependency issue",
            "schedule", "scheduling", "delayed", "not triggered",
            "cron", "job not triggered", "schedule missed",
            "pipeline not started", "pipeline delay", "late run",
            "got failed", "has failed", "failed for data date",
        ],
    },
    {
        "name": "DB Issue",
        "color": "00B0F0",
        "keywords": [
            "database", "db issue", "sql error", "timeout",
            "oracle", "postgres", "mysql", "db connectivity",
            "db timeout", "database error",
        ],
    },
    {
        "name": "Metric drop",
        "color": "7030A0",
        "keywords": [
            "metric drop", "low volume", "reduced count",
            "sudden drop", "drop in metrics", "metric decline",
            "volume drop", "low metric",
        ],
    },
    {
        "name": "Schema issue",
        "color": "FF69B4",
        "keywords": [
            "schema mismatch", "column mismatch", "header mismatch",
            "trailer mismatch", "schema issue", "schema change",
            "schema update", "schema drift", "column missing",
            "field change", "new column", "field missing",
        ],
    },
    {
        "name": "Issue Platform",
        "color": "00B0B0",
        "keywords": [
            "nifi", "airflow", "actian", "aws",
            "infrastructure issue", "server issue", "infrastructure",
            "server", "platform", "nifi issue", "airflow issue",
            "actian issue",
        ],
    },
]

OTHER_CATEGORY = {"name": "Other", "color": "D9D9D9"}
ALL_CATS = [c["name"] for c in CATEGORIES] + [OTHER_CATEGORY["name"]]
