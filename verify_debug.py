import re
from nile_pipeline_tracker.config import CATEGORIES, OTHER_CATEGORY

def score_breakdown(summary, description, issue_type="", comment="", close_notes=""):
    fields = [
        ("summary",     summary      or "", 1),
        ("description", description  or "", 1),
        ("issue_type",  issue_type   or "", 1),
        ("comment",     comment      or "", 2),
        ("close_notes", close_notes  or "", 3),
    ]
    results = {}
    for cat in CATEGORIES:
        cat_score = 0
        hits = []
        for fname, text, weight in fields:
            text_lower = text.lower()
            for kw in cat["keywords"]:
                if re.search(r"\b" + re.escape(kw.lower()) + r"\b", text_lower):
                    cat_score += weight
                    hits.append(f"{fname}[x{weight}]:{kw!r}")
        if cat_score > 0:
            results[cat["name"]] = (cat_score, hits)
    return dict(sorted(results.items(), key=lambda x: -x[1][0]))

print("=" * 70)
print("TEST 1: Pipeline failed Zero Ingested row count Unmatched")
print("Expected: Schema issue")
bd = score_breakdown(
    "Pipeline failed Zero Ingested row count Unmatched",
    "pipeline failed result shows 0 Ingested row count",
    "", "",
    "found out that there is a new column causing issues update the schema UD mapping validation",
)
for cat, (score, hits) in list(bd.items())[:5]:
    print(f"  {score:3}  {cat}")
    for h in hits:
        print(f"         {h}")

print()
print("=" * 70)
print("TEST 2: pipeline got failed due to significant data drop")
print("Expected: Configuration Issue")
bd = score_breakdown(
    "pipeline got failed due to significant data drop",
    "data metric percentage drop hits threshold",
    "",
    "Got confirmation to add parameter now pipeline is running",
    "Pipeline completed",
)
for cat, (score, hits) in list(bd.items())[:5]:
    print(f"  {score:3}  {cat}")
    for h in hits:
        print(f"         {h}")
