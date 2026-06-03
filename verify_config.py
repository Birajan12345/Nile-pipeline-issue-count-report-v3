import re
from nile_pipeline_tracker.config import CATEGORIES, OTHER_CATEGORY
from nile_pipeline_tracker.categorizer import categorize
from collections import defaultdict

# Check 1 — no cross-category keyword conflicts
kw_map = defaultdict(list)
for cat in CATEGORIES:
    for kw in cat["keywords"]:
        if "." not in kw and "*" not in kw:
            kw_map[kw.lower()].append(cat["name"])
conflicts = {k: v for k, v in kw_map.items() if len(v) > 1}
if conflicts:
    print("CONFLICTS:")
    for kw, cats in conflicts.items():
        print(f"  {kw!r} in {cats}")
else:
    print("OK: Zero keyword conflicts")

# Check 2 — keyword counts per category
print()
for cat in CATEGORIES:
    name = cat["name"]
    count = len(cat["keywords"])
    print(f"  {name:<30} {count} keywords")

# Check 3 — real ticket tests
print()
tests = [
    ("please ingest UHC Cirrus file", "", "Other"),  # "please ingest" removed (Fix 2 — too broad)
    ("files spooled and processed from nile", "", "Spooler issue"),
    ("backfill files second set not being spooled", "", "Spooler issue"),
    ("pipeline is getting failed 0 byte pgp file blocking ingestion", "", "Pipeline Schedule"),
    ("spelling error in filename misspelled file name pattern", "", "File Issue"),
    ("same sequence number in both files replacement file resent", "", "File Issue"),
    ("schema is not matching source file schema table schema", "", "Schema issue"),
    ("layout update tab names changed blank values field", "", "Schema issue"),
    ("override for metric drop data metric percentage hits threshold", "", "Metric drop"),
    ("please hold pipeline waiting on missing files hold until receive", "", "File Missing"),
    ("no longer able to find file history files missing", "", "File Missing"),
    ("disk space insufficient filegroup transaction log sqlserver", "", "DB Issue"),
    ("stuck in db writer hdfs issue hadoop issue dc-pine", "", "DB Issue"),
    ("permission denied destination copy failure pipeline stopped", "", "Pipeline Schedule"),
    ("gwf flag sftp password suspend pipeline", "", "Configuration Issue"),
    ("unmatched raw counts lineage report count does not match", "", "Unmatched raw data count"),
    ("significant data drop 0 records to be ingested causing 0 records", "", "Data Issue"),
    ("nile platform issue sftp issue ftp down connectivity issue", "", "Issue Platform"),
]
all_ok = True
for s, d, expected in tests:
    result = categorize(s, d)
    ok = result == expected
    if not ok:
        all_ok = False
    status = "OK  " if ok else "FAIL"
    print(f"  [{status}]  {s[:55]:<57} -> {result}")

print()
print("All passed!" if all_ok else "Some FAILED - review.")
