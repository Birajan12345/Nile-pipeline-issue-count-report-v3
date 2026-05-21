import re

from .config import CATEGORIES, OTHER_CATEGORY


def categorize(summary: str, description: str, issue_type: str = "") -> str:
    text = " ".join([
        summary     or "",
        description or "",
        issue_type  or "",
    ]).lower()

    best_name  = OTHER_CATEGORY["name"]
    best_score = 0

    for cat in CATEGORIES:
        score = sum(
            1 for kw in cat["keywords"]
            if re.search(r"\b" + re.escape(kw.lower()) + r"\b", text)
        )
        if score > best_score:
            best_score = score
            best_name  = cat["name"]

    return best_name
