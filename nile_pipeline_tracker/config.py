import os

CSV_PATH = os.getenv("CSV_PATH", "jira_export.csv")
OUTPUT_DIR = os.getenv("OUTPUT_DIR", ".")
REPORT_MONTH_LABEL = "April 2026"

COLUMN_ALIASES = {
    "key": ["Issue key", "Key", "Issue Key", "key"],
    "issue_id": ["Issue id", "Issue ID", "ID", "id"],
    "parent_id": ["Parent id", "Parent ID", "Parent", "Parent key"],
    "summary": ["Summary", "summary", "Title"],
    "description": ["Description", "description", "Issue Description"],
    "categories": ["Categories", "categories", "Category", "category", "GSR Category", "GSR Sub Category", "New Categories", "New Catogories", "New Category"],
    "new_categories": ["New Categories", "New Category", "GSR Sub Category", "GSR Category"],
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
            "configuration", "config", "parameter", "setup", "mapping", "env issue", "property issue",
            "config error", "configuration issue", "config change", "pipeline config",
            "misconfigured", "config update", "parameter change", "pipeline parameter",
            "wrong config",
        ],
    },
    {
        "name": "File Issue",
        "color": "4472C4",
        "keywords": [
            "corrupt file", "invalid file", "encrypted file", "damaged file", "empty file", "bad file",
            "file issue", "missing files", "file not received", "file not found", "file delay",
            "file arrival", "file drop", "source file", "file rejected", "file ingestion",
            "not ingested", "file failure",
        ],
    },
    {
        "name": "Spooler issue",
        "color": "FF0000",
        "keywords": [
            "spooler", "unscheduled", "stuck", "trigger issue", "pickup failure", "spool delay",
            "spool error", "spool failed", "spool queue", "spooling",
        ],
    },
    {
        "name": "Data Issue",
        "color": "ED7D31",
        "keywords": [
            "data mismatch", "duplicate data", "null values", "incorrect data", "missing records",
            "data issue", "data quality", "wrong data", "data discrepancy", "data error",
            "invalid data", "corrupt data", "bad data", "data not matching",
            "data validation", "data integrity", "data drop", "data count",
            "raw data", "record count", "count mismatch", "unmatched count",
        ],
    },
    {
        "name": "File Missing",
        "color": "FFC000",
        "keywords": [
            "file missing", "missing file", "no file received", "replacement file pending",
            "missing files", "missing data file",
        ],
    },
    {
        "name": "Unmatched raw data count",
        "color": "00B050",
        "keywords": [
            "count mismatch", "unmatched count", "raw count mismatch", "record mismatch",
            "unmatched raw data count", "record count mismatch", "row count mismatch",
        ],
    },
    {
        "name": "Pipeline Schedule",
        "color": "FFC000",
        "keywords": [
            "pipeline failed", "workflow failed", "cron issue", "delayed workflow", "dependency issue",
            "schedule", "scheduling", "delayed", "not triggered", "cron",
            "job not triggered", "schedule missed", "pipeline not started",
            "pipeline delay", "late run",
        ],
    },
    {
        "name": "DB Issue",
        "color": "00B0F0",
        "keywords": [
            "database", "db issue", "sql error", "timeout", "oracle", "postgres", "mysql",
            "db connectivity", "db timeout", "database error",
        ],
    },
    {
        "name": "Metric drop",
        "color": "7030A0",
        "keywords": [
            "metric drop", "low volume", "reduced count", "sudden drop", "drop in metrics",
            "metric decline", "volume drop", "low metric",
        ],
    },
    {
        "name": "Schema issue",
        "color": "FF69B4",
        "keywords": [
            "schema mismatch", "column mismatch", "header mismatch", "trailer mismatch",
            "schema issue", "schema change", "schema update", "schema drift",
            "column missing", "field change", "new column", "field missing",
        ],
    },
    {
        "name": "Issue Platform",
        "color": "00B0B0",
        "keywords": [
            "nifi", "airflow", "actian", "aws", "infrastructure issue", "server issue",
            "infrastructure", "server", "platform", "nifi issue", "airflow issue", "actian issue",
        ],
    },
]

OTHER_CATEGORY = {"name": "Other", "color": "D9D9D9"}
ALL_CATS = [c["name"] for c in CATEGORIES] + [OTHER_CATEGORY["name"]]
