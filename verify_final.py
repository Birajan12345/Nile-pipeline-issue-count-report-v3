from nile_pipeline_tracker.categorizer import categorize

tests = [
    # BMC failure → Pipeline Schedule
    ('Error when running Humana BMC pipeline trigger_bmc_workflow',
     'precheck conditions wrong Modified and published pipeline v1.2',
     '', 'Modified and published pipeline version', 'Pipeline processed successfully',
     'Pipeline Schedule'),
    # trigger_bmc failure → Pipeline Schedule
    ('trigger_bmc_workflow-uhc-unet_monthly_DataLoads_V2',
     'trigger_bmc is not recognizing these pipelines as completed',
     '', 'changed and published pipelines to get picked up', '',
     'Pipeline Schedule'),
    # Sequence file failure → File Issue
    ('WELEX History files failed 04/18/2026',
     'files failed because sequence number already used',
     '', 'We have requested client to send the replacement file', '',
     'File Issue'),
    # OLA monitoring → Other (not Config)
    ('OLA Configuration Changes and Monitoring on cigna',
     'I have made configuration changes in OLA including buffer time',
     '', 'Buffer time has been added I will monitor the OLA warnings', '',
     'Other'),
    # Config fix published → Configuration Issue (not Other)
    ('Request to Schedule dm-aetna-hrp pipeline',
     'schedule the pipeline to avoid manual trigger dependency',
     '', 'Created published and scheduld vrsion 1.7',
     'Created tested published and scheduled Version 1.7',
     'Configuration Issue'),
    # BMC update prechecks with real fix → Configuration Issue
    ('BCBS_MI Update prechecks for BMC and publish all child pipelines',
     'Update precheck with pipeline version as latest',
     '', 'BMC Pipeline prechecks updated Published and scheduled',
     'BMC Pipeline prechecks conditions updated',
     'Configuration Issue'),
    # Genuine spooler failure → Spooler issue
    ('NAS Spoolers failed in production environment',
     'Nile NAS spoolers failed pipelines got stuck',
     '', 'Joined command center bridge call ran pipelines manually',
     'NAS spoolers failed resolved',
     'Spooler issue'),
    # BMC peer review → Other (exclusion still works)
    ('Peer Review for Aetna-Trad-SCM BMC pipeline Optimization',
     'Peer review for BMC workflow pipelines verify child pipelines',
     '', 'I have reviewed the bmc pipelines',
     'Reviewed and scheduled',
     'Other'),
    # Duplicate → Other (exclusion still works)
    ('Unmatched mbr enrollment raw data count pipeline failed',
     'pipeline got failed due to unmatched record',
     '', 'There is another ticket already opened', '',
     'Other'),
    # NAS spooler + duplicate tag → Other
    ('Nile NAS spooler failures',
     'Nile NAS spoolers failed pipelines got stuck HUE browser',
     '', 'Duplicate ticket created', '',
     'Other'),
]

print('Running 10 verification tests...')
print()
all_ok = True
for s, d, it, c, cn, expected in tests:
    result = categorize(s, d, it, c, cn)
    ok = result == expected
    if not ok:
        all_ok = False
    status = 'OK  ' if ok else 'FAIL'
    print(f'  [{status}]  {s[:52]:<54} -> {result}')

print()
print('All 10 tests passed!' if all_ok else 'Some FAILED — review changes.')
