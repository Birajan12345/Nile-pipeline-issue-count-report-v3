import re
from nile_pipeline_tracker.categorizer import EXCLUSION_PATTERNS
from nile_pipeline_tracker.config import CATEGORIES

comment = "Data metric percentage drop hits threshold forwarded email"
cn = ""
exclusion_text = (cn + " " + comment).lower()
print("Exclusion text:", repr(exclusion_text))
print()
matched = [p for p in EXCLUSION_PATTERNS if p.lower() in exclusion_text]
print("Exclusion patterns matched:", matched if matched else "NONE — no exclusion fired")
print()
print("Scoring breakdown (comment only, weight 2):")
for cat in CATEGORIES:
    hits = []
    for kw in cat["keywords"]:
        if re.search(r"\b" + re.escape(kw.lower()) + r"\b", comment.lower()):
            hits.append(kw)
    if hits:
        score = len(hits) * 2
        print(f"  {cat['name']:<30} score={score}  hits={hits}")
