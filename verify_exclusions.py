from nile_pipeline_tracker.categorizer import categorize

tests = [
    # Duplicate → Other
    ("Unmatched mbr enrollment raw data count pipeline got failed",
     "pipeline failed due to unmatched record",
     "", "There is another ticket already opened", "",
     "Other"),
    # Client side → Other
    ("OSCAR no claim volume Importance High",
     "no claim volume from client",
     "", "I realized this was on the client side", "",
     "Other"),
    # Spooler admin rename → Other
    ("Update spooler name prod",
     "update the spooler name in prod",
     "", "Hi updated the spooler name Thanks", "Updated Spooler name",
     "Other"),
    # BMC review → Other
    ("BMC Pipelines Review and Precheck Verification list for 15 clients",
     "Review all BMC pipelines for 20 clients",
     "", "Listed all the required BMC pipeline list", "Listed all the bmc optimization pipelines",
     "Other"),
    # Genuine spooler failure → NOT Other
    ("NAS Spoolers failed in production environment",
     "Nile NAS spoolers failed pipelines got stuck",
     "", "Joined command center bridge call", "NAS spoolers failed",
     "Spooler issue"),
    # Genuine schema issue → NOT Other
    ("uhc cirrus pipeline failed due to schema mismatch",
     "Source file schema is not matching with the Table Schema",
     "", "", "tasks completed schema updated",
     "Schema issue"),
]

all_ok = True
for s, d, it, c, cn, expected in tests:
    result = categorize(s, d, it, c, cn)
    ok = result == expected
    if not ok:
        all_ok = False
    status = "OK  " if ok else "FAIL"
    print(f"  [{status}]  {s[:50]:<52} -> {result}")

print()
print("All passed!" if all_ok else "Some FAILED.")
