import re

from .config import CATEGORIES, OTHER_CATEGORY


def categorize(summary: str, description: str, issue_type: str = "") -> str:
    text = " ".join([summary, description or "", issue_type or ""]).lower()
    for cat in CATEGORIES:
        for kw in cat["keywords"]:
            if re.search(r"\b" + re.escape(kw.lower()) + r"\b", text):
                return cat["name"]
    return OTHER_CATEGORY["name"]
