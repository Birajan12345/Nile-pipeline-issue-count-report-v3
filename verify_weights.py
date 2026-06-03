from nile_pipeline_tracker.categorizer import categorize

tests = [
    # Schema fix should win over data/metric symptom
    (
        "Pipeline failed Zero Ingested row count Unmatched",
        "pipeline failed result shows 0 Ingested row count",
        "", "",
        "found out that there is a new column causing issues update the schema UD mapping validation",
        "Schema issue",
    ),
    # Config fix should win over metric symptom
    (
        "pipeline got failed due to significant data drop",
        "data metric percentage drop hits threshold",
        "",
        "Got confirmation to add parameter now pipeline is running",
        "Pipeline completed",
        "Configuration Issue",
    ),
    # True metric drop (no resolution context overriding)
    (
        "FWAM BCBS North Dakota CDF File Record Counts Low for 2/26 Run",
        "record counts very low only getting 25 percent of normal CDF records",
        "",
        "ran all loads fixing and investigating all member discards",
        "",
        "Metric drop",
    ),
    # Spooler failure (not just ingestion request)
    (
        "NAS Spoolers failed in production environment",
        "Nile NAS spoolers failed pipelines got stuck HUE browser inaccessible",
        "",
        "joined command center bridge call ran pipelines manually",
        "Nile NAS spoolers failed",
        "Spooler issue",
    ),
    # Schema should win
    (
        "uhc-cirrus pipeline failed due to schema mismatch",
        "Source file schema is not matching with the Table Schema check Column names Data-types",
        "", "",
        "tasks completed schema updated",
        "Schema issue",
    ),
]

all_ok = True
for s, d, it, c, cn, expected in tests:
    result = categorize(s, d, it, c, cn)
    ok = result == expected
    if not ok:
        all_ok = False
    status = "OK  " if ok else "FAIL"
    print(f"  [{status}]  {s[:55]:<57} -> {result}")

print()
print("All passed!" if all_ok else "Some FAILED — review weights.")
