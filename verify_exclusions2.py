from nile_pipeline_tracker.categorizer import categorize

tests = [
    # OLA monitoring → Other
    ("OLA Configuration Changes and Monitoring on cigna",
     "I have made configuration changes in OLA including adding buffer time",
     "", "Buffer time has been added and I will monitor the OLA warnings", "",
     "Other"),
    # Email notification setup → Other
    ("Request to add EDMs distro list to NILE email notifications",
     "Please add EDMs distribution list to NILE ingestion report",
     "", "Data metric percentage drop hits threshold forwarded email", "",
     "Other"),
    # BMC setup → Other
    ("Please fix Kaiser-wa BMC pipeline",
     "please fix trigger_bmc_workflow-kaiser_wa_monthly as it is not schedule",
     "", "Fixed pre-check publish child pipelines and scheduled",
     "Fixed pre-check publish child pipelines and scheduled",
     "Other"),
    # MCL run request → Other
    ("Please Run These MCL Quarterly Pipelines",
     "Please run the MCL quarterly pipelines",
     "", "Successfully loaded all the files to target tables", "",
     "Other"),
    # Genuine spooler failure — NOT Other
    ("NAS Spoolers failed in production environment",
     "Nile NAS spoolers failed pipelines got stuck HDFS NameNode safe mode",
     "", "Joined command center bridge call ran pipelines manually",
     "Nile NAS spoolers failed",
     "Spooler issue"),
    # Genuine schema issue — NOT Other
    ("uhc-cirrus pipeline failed due to schema mismatch",
     "Source file schema is not matching with the Table Schema",
     "", "", "tasks completed schema updated",
     "Schema issue"),
    # Molina data issue — should NOT be excluded
    ("Molina Raw Data Needed data discrepancies CCV audits",
     "Please provide raw data for 5 CCV inpatient files data discrepancies",
     "", "Shared raw files for analysis", "",
     "Data Issue"),
    # BCBS-ND blank values — schema, NOT Other
    ("BCBS-ND File With Blank Values EPIC_CD RMS Field",
     "blank values in EPIC_CD field of weekly overpayment files",
     "", "Please open ticket with RCA engineering team to investigate", "",
     "Schema issue"),
]

all_ok = True
for s, d, it, c, cn, expected in tests:
    result = categorize(s, d, it, c, cn)
    ok = result == expected
    if not ok:
        all_ok = False
    status = "OK  " if ok else "FAIL"
    print(f"  [{status}]  {s[:52]:<54} -> {result}")

print()
print("All passed!" if all_ok else "Some FAILED — review patterns.")
