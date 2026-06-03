import re

from .config import CATEGORIES, OTHER_CATEGORY

EXCLUSION_PATTERNS = [
    # Duplicate tickets
    "duplicate ticket", "not needed. duplicate", "duplicate",
    "closing out ticket. ticket not needed",
    "there is already another ticket",
    "another ticket already opened",
    # Client-side / no Nile action
    "client side and shouldn't have created",
    "on the client side",
    "this is on the client",
    "not a nile issue",
    "this is not with nile",
    "not sure yet how we should proceed",
    "no action required from nile",
    # Admin / setup / monitoring tasks (resolved by setup not fix)
    "removed the unused spoolers",
    "updated spooler name",
    "removed the pipeline and spoolers from the nile application",
    "both spooler and pipeline scheduled",
    "spoolers and pipelines were rescheduled",
    "moved elist spooler to event based",
    "listed all the bmc optimization pipelines",
    "listed all the required bmc pipeline",
    "i have provided list of bmc optimization",
    "generated files to finance",
    "shared raw, gdf and cdf details",
    # Email/notification setup
    "add edm's distribution list",
    "adding 2 email list",
    "adding 2 email",
    "add edm_rca",
    "distribution list of new platform",
    "edm distro",
    "edm notification",
    "add edm",
    # OLA monitoring — buffer time setup tickets
    "i will monitor a few executions",
    "buffer time has been added",
    "i will monitor the ola warnings",
    "ola monitoring are not valid",
    "ola configuration changes and monitoring",
    "will monitor one more day",
    # Spooler admin tasks
    "created a new spooler and new pipeline",
    "created new spoolers and new pipeline",
    "move elist spooler",
    "moved elist spooler",
    "update this spooler",
    "update spooler",
    # BMC / pipeline setup and scheduling
    "updating the pre-checks",
    "updated all prechecks",
    "need to create 2 new bmc pipelines",
    "please fix trigger_bmc",
    "fixed pre-check, publish child pipelines",
    "fixed pre-check publish child pipelines",   # no-comma variant
    "publish child pipelines",
    # MCL / table run requests
    "please run the mcl",
    "run the mcl",
    "please run these mcl",
    "successfully loaded all the files to target tables",
    "files loaded successfully and user confirmed",
    # Table update requests
    "please do a full replacement of the existing table",
    "full replacement of the existing table",
    # Data requests (no pipeline failure)
    "provided the requestor's analysis",
    "provided the necessary analysis for end user",
    # HDFS file copy/move
    "files moved and updated permissions",
    "files copied please verify",
    "files copied",
    # Retro / ARD jobs
    "ard-prod", "retro process",
    # Test/dummy
    "testing. not needed",
    "not needed. thank you",
]


def categorize(summary: str, description: str, issue_type: str = "",
               comment: str = "", close_notes: str = "") -> str:

    # Check exclusion patterns in Close Notes, Comment, and Summary
    # If matched → return Other immediately (admin/duplicate/non-issue)
    exclusion_text = " ".join([
        (close_notes or "").lower(),
        (comment     or "").lower(),
        (summary     or "").lower(),
    ])
    for pattern in EXCLUSION_PATTERNS:
        if pattern.lower() in exclusion_text:
            return OTHER_CATEGORY["name"]

    # Build weighted text — resolution fields score higher
    fields = [
        (summary      or "", 1),
        (description  or "", 1),
        (issue_type   or "", 1),
        (comment      or "", 2),   # 2x weight
        (close_notes  or "", 3),   # 3x weight — root cause
    ]

    best_name  = OTHER_CATEGORY["name"]
    best_score = 0

    for cat in CATEGORIES:
        score = 0
        for text, weight in fields:
            text_lower = text.lower()
            for kw in cat["keywords"]:
                if re.search(r"\b" + re.escape(kw.lower()) + r"\b", text_lower):
                    score += weight
        if score > best_score:
            best_score = score
            best_name  = cat["name"]

    return best_name
