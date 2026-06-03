import csv
import os
import re
import sys
from collections import Counter
from datetime import datetime

from .config import COLUMN_ALIASES, HTML_FIELD_LABELS
from .categorizer import categorize


def normalize_text(raw: str) -> str:
    if raw is None:
        return ""
    text = str(raw)
    text = re.sub(r"[\r\n\t]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def detect_column(headers: list, aliases: list) -> str | None:
    """Return the actual CSV header that matches any alias."""
    headers_lower = {h.lower(): h for h in headers}
    for alias in aliases:
        if alias.lower() in headers_lower:
            return headers_lower[alias.lower()]
    return None


def read_csv(filepath: str) -> tuple[list[dict], dict]:
    """Read Jira CSV and return (rows, column_map)."""
    if not os.path.exists(filepath):
        print(f"\n  ERROR: CSV file not found: {filepath}")
        print("  Please export from Jira and update CSV_PATH in the script.")
        sys.exit(1)

    with open(filepath, newline="", encoding="latin-1") as f:
        reader = csv.DictReader(f)
        headers = reader.fieldnames or []
        col_map = {
            field: detect_column(headers, aliases)
            for field, aliases in COLUMN_ALIASES.items()
        }

        raw_rows = []
        for row in reader:
            cleaned = {
                field: normalize_text(row.get(col, "")) if col else ""
                for field, col in col_map.items()
            }
            raw_rows.append(cleaned)

    print(f"  CSV loaded: {len(raw_rows)} rows, {len(headers)} columns")
    missing = [f for f, c in col_map.items() if c is None and f in ("key", "summary", "created")]
    if missing:
        print(f"  WARNING: Could not find columns: {missing}")
        print(f"  Available headers: {headers}")

    return raw_rows, col_map


def _parse_date(raw: str) -> str:
    """Try common Jira and HTML date formats and return YYYY-MM-DD."""
    raw = normalize_text(raw)
    raw = re.sub(r"\s+(EDT|EST|CST|CDT|PST|PDT)$", "", raw, flags=re.I)
    raw = raw.replace("&nbsp;", " ")
    formats = [
        "%d/%b/%y %I:%M %p",
        "%d/%b/%Y %I:%M %p",
        "%d/%b/%y",
        "%d/%b/%Y",
        "%Y-%m-%dT%H:%M:%S.%f%z",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
        "%d-%b-%Y",
        "%m/%d/%Y",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(raw.strip(), fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return raw[:10] if raw else ""


def _parse_subtext(subtext: str) -> tuple[str, str, str]:
    subtext = normalize_text(subtext).replace("\xa0", " ")
    created = ""
    updated = ""
    resolved = ""

    for label in ("Created", "Updated", "Resolved"):
        match = re.search(rf"{label}:\s*([^ ](?:.*?))(?=(?:Created:|Updated:|Resolved:|$))", subtext, flags=re.I)
        if match:
            value = normalize_text(match.group(1))
            if label == "Created":
                created = _parse_date(value)
            elif label == "Updated":
                updated = _parse_date(value)
            elif label == "Resolved":
                resolved = _parse_date(value)
    return created, updated, resolved


def _canonical_html_label(label: str) -> str | None:
    normalized = normalize_text(label).lower().rstrip(":")
    for canonical, aliases in HTML_FIELD_LABELS.items():
        for alias in aliases:
            if normalized == normalize_text(alias).lower().rstrip(":"):
                return canonical
    return None


def _extract_issue_links(scope, fields: dict):
    inward_cloners = []
    outward_cloners = []
    inward_relates = []

    targets = []
    if isinstance(scope, list):
        targets = scope
    else:
        targets = [scope]

    for context in targets:
        for link_table in context.select("table"):
            if not link_table.get_text(" ", strip=True):
                continue
            if "Issue Links" not in link_table.get_text(" "):
                continue

            for row in link_table.select("tr"):
                cells = row.find_all("td", recursive=False)
                if len(cells) < 2:
                    continue

                relation = normalize_text(cells[0].get_text(" ", strip=True)).lower()
                target = ""
                anchor = cells[1].find("a")
                if anchor:
                    target = normalize_text(anchor.get_text(" ", strip=True))
                if not target:
                    continue

                if "is cloned by" in relation or "cloned by" in relation:
                    inward_cloners.append(target)
                elif "clones" in relation:
                    outward_cloners.append(target)
                elif "is related by" in relation or "is related to" in relation:
                    inward_relates.append(target)
                elif "relates" in relation:
                    inward_relates.append(target)

    fields["Inward issue link (Cloners)"] = ", ".join(inward_cloners)
    fields["Outward issue link (Cloners)"] = ", ".join(outward_cloners)
    fields["Inward issue link (Relates)"] = ", ".join(inward_relates)


def _read_html_fields(scope, fields: dict):
    for row in scope.select("tr"):
        cells = row.find_all("td", recursive=False)
        if len(cells) < 2:
            continue

        label = normalize_text(cells[0].get_text(" ", strip=True)).rstrip(":")
        value = normalize_text(cells[1].get_text(" ", strip=True))
        if not label:
            continue

        canonical = _canonical_html_label(label)
        if not canonical:
            continue

        if canonical == "categories":
            if "new" in label.lower() or "sub" in label.lower():
                fields["New Categories"] = value
            elif not fields.get("Categories"):
                fields["Categories"] = value
            else:
                fields["Categories"] = f"{fields['Categories']}, {value}" if value else fields["Categories"]
        elif canonical == "new_categories":
            fields["New Categories"] = value
        else:
            if fields.get(canonical) and canonical in ("comments", "close_notes"):
                fields[canonical] = f"{fields[canonical]}\n\n{value}" if value else fields[canonical]
            else:
                fields[canonical] = value


def _gather_comment_blocks(scope):
    comments = []
    for block in scope.select("tr[id^='comment-body'] td"):
        comment_text = normalize_text(block.get_text(" ", strip=True))
        if comment_text:
            comments.append(comment_text)
    return comments


def _find_parent_id(scope) -> str:
    parent = scope.select_one("#parent_issue_summary")
    if parent and parent.has_attr("data-issue-key"):
        return normalize_text(parent["data-issue-key"])
    if parent:
        return normalize_text(parent.get_text(" ", strip=True))
    return ""


def _collect_issue_block(header_table):
    block = []
    current = header_table.find_next_sibling()
    while current:
        if current.name == "table" and "tableBorder" in (current.get("class") or []) and current.find("h3", class_="formtitle"):
            break
        block.append(current)
        if current.name == "hr" and "fullcontent" in (current.get("class") or []):
            break
        current = current.find_next_sibling()
    return block


def _extract_html_issue_rows(soup) -> list[dict]:
    rows = []
    for header_table in soup.select("table.tableBorder"):
        title = header_table.find("h3", class_="formtitle")
        if not title:
            continue

        # Plain <a href> link (no issue-link class in Service Desk HTML export)
        anchor = title.find("a")
        if anchor:
            summary = normalize_text(anchor.get_text(strip=True))
            href = anchor.get("href", "")
            m = re.search(r"/browse/([A-Z]+-\d+)", href)
            issue_key = m.group(1) if m else ""
        else:
            title_text = normalize_text(title.get_text(" ", strip=True))
            m = re.search(r"\[([A-Z]+-\d+)\]", title_text)
            issue_key = m.group(1) if m else ""
            summary = re.sub(r"^\s*\[[A-Z]+-\d+\]\s*", "", title_text)
        # Strip any trailing date metadata that leaked from the subText span
        summary = re.sub(r"\s+Created:\s.*$", "", summary, flags=re.I).strip()

        subtext = title.find("span", class_="subText")
        created, updated, _ = _parse_subtext(subtext.get_text(" ", strip=True)) if subtext else ("", "", "")

        issue_id = issue_key or ""
        if issue_key and issue_id.isupper():
            issue_id = issue_key

        fields = {
            "key": issue_key,
            "issue_id": issue_id,
            "parent_id": "",
            "summary": summary,
            "description": "",
            "Comment": "",
            "close_notes": "",
            "Categories": "",
            "New Categories": "",
            "Assign": "",
            "status": "",
            "created": created,
            "updated": updated,
            "assignment_group": "",
            "Assignee": "",
            "Reporter": "",
            "resolution": "",
            "comments": "",
            "Inward issue link (Cloners)": "",
            "Outward issue link (Cloners)": "",
            "Inward issue link (Relates)": "",
            "Labels": "",
            "issue_type": "",
            "priority": "",
        }

        # Status is in the header table itself, not in the block below it
        for tr in header_table.select("tr"):
            cells = tr.find_all("td", recursive=False)
            if len(cells) >= 2:
                lab = normalize_text(cells[0].get_text(" ", strip=True)).rstrip(":")
                if lab.lower() == "status":
                    fields["status"] = normalize_text(cells[1].get_text(" ", strip=True))
                    break

        block = _collect_issue_block(header_table)
        comment_texts = []

        for node in block:
            if not getattr(node, "find_all", None):
                continue
            _read_html_fields(node, fields)
            if fields.get("description") == "":
                description_node = node.select_one("td#descriptionArea")
                if description_node:
                    fields["description"] = normalize_text(description_node.get_text(" ", strip=True))
            comment_texts.extend(_gather_comment_blocks(node))
            if not fields["parent_id"]:
                fields["parent_id"] = _find_parent_id(node)
            if not fields["Assign"] and fields.get("assignee"):
                fields["Assign"] = fields.get("assignee")

        _extract_issue_links(block, fields)

        fields["Comment"] = "\n\n".join(comment_texts)
        fields["comments"] = fields["Comment"]
        fields["Custom field (Comments)"] = fields["Comment"]
        fields["Custom field (Assignment Group)"] = fields.get("assignment_group", "")
        fields["Review Required"] = "Yes" if fields.get("Category", "") == "Other" else ""
        fields["Assign"] = fields.get("Assignee", fields.get("Assign", ""))

        rows.append(fields)
    return rows


def read_html(filepath: str) -> tuple[list[dict], dict]:
    if not os.path.exists(filepath):
        print(f"\n  ERROR: HTML file not found: {filepath}")
        sys.exit(1)

    try:
        from bs4 import BeautifulSoup
    except ImportError:
        print("\n  ERROR: beautifulsoup4 and lxml are required to parse HTML.")
        print("  Install with: pip install beautifulsoup4 lxml")
        sys.exit(1)

    with open(filepath, encoding="utf-8", errors="replace") as f:
        soup = BeautifulSoup(f, "lxml")

    raw_rows = _extract_html_issue_rows(soup)
    print(f"  HTML loaded: {len(raw_rows)} issues parsed")
    return raw_rows, {}


def parse_rows(raw_rows: list) -> list[dict]:
    rows = []
    for r in raw_rows:
        # raw_rows already have canonical field names as keys (both CSV and HTML paths)
        def get(field: str) -> str:
            return normalize_text(r.get(field, ""))

        summary = get("summary")
        description = get("description")
        # Comments and categories may be stored under canonical or HTML-style keys
        comments = get("comments") or normalize_text(r.get("Comment", ""))
        close_notes = get("close_notes")
        resolution = get("resolution")
        assignment_group = get("assignment_group")
        categories = get("categories") or normalize_text(r.get("Categories", ""))
        new_categories = get("new_categories") or normalize_text(r.get("New Categories", ""))
        assignee = get("assignee") or get("assign") or normalize_text(r.get("Assign", ""))
        issue_type = get("issue_type")

        category = categorize(
            summary,
            description,
            issue_type,
            comment=comments,
            close_notes=close_notes,
        )

        rows.append({
            "Key": get("key"),
            "Issue id": get("issue_id") or get("key"),
            "Parent id": get("parent_id"),
            "Summary": summary,
            "Description": description,
            "Comment": comments,
            "Close Notes": close_notes,
            "Categories": categories,
            "New Categories": new_categories,
            "Assign": assignee,
            "Status": get("status"),
            "Created": _parse_date(get("created")),
            "Updated": _parse_date(get("updated")),
            "Custom field (Assignment Group)": assignment_group,
            "Assignee": assignee or "Unassigned",
            "Reporter": get("reporter"),
            "Resolution": resolution,
            "Custom field (Comments)": comments,
            "Inward issue link (Cloners)": get("inward_cloners") or normalize_text(r.get("Inward issue link (Cloners)", "")),
            "Outward issue link (Cloners)": get("outward_cloners") or normalize_text(r.get("Outward issue link (Cloners)", "")),
            "Inward issue link (Relates)": get("inward_relates") or normalize_text(r.get("Inward issue link (Relates)", "")),
            "Category": category,
            "Review Required": "Yes" if category == "Other" else "",
        })
    return rows


def detect_month_label(rows: list) -> str:
    """Infer month label from the most common creation month in data."""
    months = Counter()
    for r in rows:
        d = r.get("Created", "")
        if len(d) >= 7:
            months[d[:7]] += 1
    if months:
        top = months.most_common(1)[0][0]
        try:
            return datetime.strptime(top, "%Y-%m").strftime("%B %Y")
        except ValueError:
            return top
    return "Report"


def read_xlsx(filepath: str) -> tuple[list[dict], dict]:
    """Read a Jira XLSX export and return (rows, column_map)."""
    try:
        import openpyxl
    except ImportError:
        print("\n  ERROR: openpyxl is required to read .xlsx files.")
        print("  Install with: pip install openpyxl")
        sys.exit(1)

    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)

    try:
        header_row = next(rows_iter)
    except StopIteration:
        print(f"\n  ERROR: XLSX file appears empty: {filepath}")
        sys.exit(1)

    headers = [str(h).strip() if h is not None else "" for h in header_row]
    col_map = {
        field: detect_column(headers, aliases)
        for field, aliases in COLUMN_ALIASES.items()
    }
    field_idx = {
        field: headers.index(col)
        for field, col in col_map.items()
        if col and col in headers
    }

    raw_rows = []
    for row in rows_iter:
        if all(cell is None for cell in row):
            continue
        cleaned = {}
        for field in col_map:
            if field in field_idx:
                idx = field_idx[field]
                val = row[idx] if idx < len(row) else None
                cleaned[field] = normalize_text(str(val) if val is not None else "")
            else:
                cleaned[field] = ""
        raw_rows.append(cleaned)

    wb.close()
    print(f"  XLSX loaded: {len(raw_rows)} rows, {len(headers)} columns")
    missing = [f for f, c in col_map.items()
               if c is None and f in ("key", "summary", "created")]
    if missing:
        print(f"  WARNING: Could not find columns: {missing}")

    return raw_rows, col_map
