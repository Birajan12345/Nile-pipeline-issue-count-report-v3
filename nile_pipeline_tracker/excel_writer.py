from collections import defaultdict

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter

from .config import CATEGORIES, OTHER_CATEGORY, ALL_CATS

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(color="FFFFFF", bold=True, size=11)
ALT_FILL = PatternFill("solid", fgColor="F2F7FF")
THIN = Side(style="thin", color="CCCCCC")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CAT_COLOR = {c["name"]: c["color"] for c in CATEGORIES}
CAT_COLOR[OTHER_CATEGORY["name"]] = OTHER_CATEGORY["color"]


def load_prev_month_counts(output_dir: str) -> dict:
    """
    Scan output/ for the most recent CSV report from a prior month.
    Returns a dict of {category_name: count} or empty dict if none found.
    """
    import csv as _csv
    import glob
    import os
    import time

    pattern   = os.path.join(output_dir, "Nile_Pipeline_Issue_Count_*.csv")
    files     = sorted(glob.glob(pattern), key=os.path.getmtime, reverse=True)

    if not files:
        return {}

    now       = time.time()
    prev_file = None
    for f in files:
        if now - os.path.getmtime(f) > 600:
            prev_file = f
            break

    if not prev_file:
        return {}

    counts = {}
    try:
        with open(prev_file, newline="", encoding="utf-8") as f:
            reader = _csv.DictReader(f)
            for row in reader:
                cat = row.get("Category", "").strip()
                if cat:
                    counts[cat] = counts.get(cat, 0) + 1
    except Exception:
        return {}

    return counts


def write_cover_sheet(ws, rows: list[dict], month_label: str,
                      cat_counts: dict, prev_counts: dict):
    """
    Executive cover sheet — the first sheet the report opens on.
    Contains: branded header, 4 KPI cards, health status,
    category snapshot table with trend arrows and mini bars.
    All data comes from rows/cat_counts — no manual input needed.
    """
    ws.title = "Cover Sheet"

    total       = len(rows)
    other_count = cat_counts.get(OTHER_CATEGORY["name"], 0)
    other_pct   = round(other_count / total * 100, 1) if total else 0

    # Find top 2 categories (excluding Other)
    issue_cats = [(n, c) for n, c in cat_counts.items()
                  if n != OTHER_CATEGORY["name"] and c > 0]
    issue_cats.sort(key=lambda x: x[1], reverse=True)
    top1_name  = issue_cats[0][0] if issue_cats else "N/A"
    top1_count = issue_cats[0][1] if issue_cats else 0
    top2_name  = issue_cats[1][0] if len(issue_cats) > 1 else ""
    top2_count = issue_cats[1][1] if len(issue_cats) > 1 else 0

    # ── Branded header rows 1-3 ────────────────────────────────
    NAVY  = "1F3864"
    WHITE = "FFFFFF"
    LIGHT = "B5D4F4"

    ws.merge_cells("A1:H1")
    h1 = ws["A1"]
    h1.value     = "NILE Platform — Production Support"
    h1.font      = Font(bold=True, size=16, color=WHITE)
    h1.fill      = PatternFill("solid", fgColor=NAVY)
    h1.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:H2")
    h2 = ws["A2"]
    h2.value     = f"Monthly Issue Report  ·  {month_label}  ·  Auto-generated"
    h2.font      = Font(size=11, color=LIGHT)
    h2.fill      = PatternFill("solid", fgColor=NAVY)
    h2.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 22

    ws.merge_cells("A3:H3")
    ws["A3"].fill = PatternFill("solid", fgColor=NAVY)
    ws.row_dimensions[3].height = 8

    # ── 4 KPI cards (rows 5-6) ─────────────────────────────────
    kpis = [
        ("Total Issues",   str(total),       "E6F1FB", "0C447C", "185FA5"),
        (top1_name,        str(top1_count),  "FCEBEB", "791F1F", "A32D2D"),
        (top2_name or "—", str(top2_count),  "FAEEDA", "633806", "854F0B"),
        ("Needs Review",   f"{other_pct}%",  "FCEBEB", "791F1F", "A32D2D"),
    ]

    ws.row_dimensions[4].height = 10
    ws.row_dimensions[5].height = 20
    ws.row_dimensions[6].height = 32
    ws.row_dimensions[7].height = 10

    for ki, (label, value, bg, txt_dark, txt_mid) in enumerate(kpis):
        col_start    = ki * 2 + 1
        col_end      = col_start + 1
        col_letter_s = get_column_letter(col_start)
        col_letter_e = get_column_letter(col_end)

        ws.merge_cells(f"{col_letter_s}5:{col_letter_e}5")
        lc = ws[f"{col_letter_s}5"]
        lc.value     = label
        lc.font      = Font(size=10, color=txt_mid)
        lc.fill      = PatternFill("solid", fgColor=bg)
        lc.alignment = Alignment(horizontal="center", vertical="center")
        lc.border    = Border(
            left=Side(style="thin", color=txt_mid),
            right=Side(style="thin", color=txt_mid),
            top=Side(style="thin", color=txt_mid),
        )

        ws.merge_cells(f"{col_letter_s}6:{col_letter_e}6")
        vc = ws[f"{col_letter_s}6"]
        vc.value     = value
        vc.font      = Font(bold=True, size=22, color=txt_dark)
        vc.fill      = PatternFill("solid", fgColor=bg)
        vc.alignment = Alignment(horizontal="center", vertical="center")
        vc.border    = Border(
            left=Side(style="thin", color=txt_mid),
            right=Side(style="thin", color=txt_mid),
            bottom=Side(style="thin", color=txt_mid),
        )

    # ── Health status row 9 ────────────────────────────────────
    ws.row_dimensions[8].height  = 10
    ws.row_dimensions[9].height  = 28
    ws.row_dimensions[10].height = 10

    if other_pct > 40:
        health_bg   = "FFF3CD"
        health_txt  = "856404"
        health_icon = "⚠"
        health_msg  = (
            f"Needs Review rate is high ({other_pct}%). "
            f"Top issue: {top1_name} ({top1_count}). "
            "Consider expanding keywords in config.py."
        )
    else:
        health_bg   = "D4EDDA"
        health_txt  = "155724"
        health_icon = "✓"
        health_msg  = (
            f"Categorization rate is healthy. "
            f"Top issue this month: {top1_name} ({top1_count}). "
            f"Total issues tracked: {total}."
        )

    ws.merge_cells("A9:H9")
    hc = ws["A9"]
    hc.value     = f"{health_icon}  Health Status:  {health_msg}"
    hc.font      = Font(size=11, color=health_txt, bold=True)
    hc.fill      = PatternFill("solid", fgColor=health_bg)
    hc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    hc.border    = Border(
        left=Side(style="thin", color=health_txt),
        right=Side(style="thin", color=health_txt),
        top=Side(style="thin", color=health_txt),
        bottom=Side(style="thin", color=health_txt),
    )

    # ── Category snapshot table from row 12 ────────────────────
    ws.row_dimensions[11].height = 10

    snap_hdr_fill = PatternFill("solid", fgColor=NAVY)
    snap_hdrs     = ["Category", "Count", "% Share", "vs Last Month", "Trend Bar"]

    ws.row_dimensions[12].height = 22
    for ci, h in enumerate(snap_hdrs, 1):
        c = ws.cell(row=12, column=ci, value=h)
        c.fill      = snap_hdr_fill
        c.font      = Font(bold=True, size=11, color=WHITE)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = BOX

    sorted_cats = [(n, cat_counts[n]) for n in ALL_CATS if cat_counts.get(n, 0) > 0]
    max_count   = max((c for _, c in sorted_cats), default=1)

    for ri, (cat_name, count) in enumerate(sorted_cats, start=13):
        pct      = round(count / total * 100, 1) if total else 0
        color    = CAT_COLOR.get(cat_name, "D9D9D9")
        alt      = ri % 2 == 0
        alt_fill = PatternFill("solid", fgColor="F2F7FF") if alt else None

        prev = prev_counts.get(cat_name, 0)
        diff = count - prev
        if prev == 0 and count > 0:
            vs_txt  = "New"
            vs_font = Font(size=11, color="185FA5", bold=True)
        elif diff > 0:
            vs_txt  = f"▲ +{diff}"
            vs_font = Font(size=11, color="3B6D11", bold=True)
        elif diff < 0:
            vs_txt  = f"▼ {diff}"
            vs_font = Font(size=11, color="A32D2D", bold=True)
        else:
            vs_txt  = "— same"
            vs_font = Font(size=11, color="888780")

        bar_len = max(1, round((count / max_count) * 20))
        bar_txt = "█" * bar_len

        row_data = [cat_name, count, f"{pct}%", vs_txt, bar_txt]
        for ci, val in enumerate(row_data, 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.border = BOX
            if alt_fill:
                c.fill = alt_fill
            if ci == 1:
                c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            elif ci == 5:
                c.font      = Font(size=10, color=color, bold=True)
                c.alignment = Alignment(horizontal="left", vertical="center")
            elif ci == 4:
                c.font      = vs_font
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="center", vertical="center")
                if ci == 2:
                    c.font = Font(bold=True, size=11)

        ws.row_dimensions[ri].height = 20

    # Total row
    total_ri = 13 + len(sorted_cats)
    tot_fill = PatternFill("solid", fgColor="E8EDF5")
    for ci in range(1, 6):
        c = ws.cell(row=total_ri, column=ci)
        c.fill      = tot_fill
        c.border    = BOX
        c.font      = Font(bold=True, size=11)
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=total_ri, column=1, value="TOTAL").alignment = \
        Alignment(horizontal="left", vertical="center", indent=1)
    ws.cell(row=total_ri, column=2, value=total)
    ws.cell(row=total_ri, column=3, value="100%")
    ws.row_dimensions[total_ri].height = 22

    # Footer note
    footer_row = total_ri + 2
    ws.merge_cells(f"A{footer_row}:H{footer_row}")
    fc = ws[f"A{footer_row}"]
    fc.value = (
        "All figures are auto-generated from the Jira export. "
        "Re-run the tracker to refresh. "
        "Needs Review tickets require manual categorization."
    )
    fc.font      = Font(size=9, color="888780", italic=True)
    fc.alignment = Alignment(horizontal="center")
    ws.row_dimensions[footer_row].height = 16

    # Column widths
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 24
    ws.column_dimensions["F"].width = 10
    ws.column_dimensions["G"].width = 10
    ws.column_dimensions["H"].width = 10

    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A12"


def write_raw_sheet(ws, rows: list[dict], month_label: str):
    ws.title = "Raw Tickets"
    cols = [
        "Key", "Issue id", "Parent id", "Summary", "Description", "Comment", "Close Notes",
        "Categories", "New Categories", "Assign", "Status", "Created", "Updated",
        "Custom field (Assignment Group)", "Assignee", "Reporter", "Resolution",
        "Custom field (Comments)", "Inward issue link (Cloners)", "Outward issue link (Cloners)",
        "Inward issue link (Relates)", "Category", "Review Required",
    ]

    ws.merge_cells(f"A1:{get_column_letter(len(cols))}1")
    title = ws["A1"]
    title.value = f"NILE Pipeline Issues — {month_label}"
    title.font = Font(bold=True, size=14, color="1F3864")
    title.alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 28

    for ci, col_name in enumerate(cols, 1):
        c = ws.cell(row=2, column=ci, value=col_name)
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BOX
    ws.row_dimensions[2].height = 22

    for ri, row in enumerate(rows, 3):
        for ci, col_name in enumerate(cols, 1):
            val = row.get(col_name, "")
            wrap = col_name in {"Summary", "Description", "Comment", "Close Notes", "Custom field (Comments)"}
            c = ws.cell(row=ri, column=ci, value=val)
            c.alignment = Alignment(wrap_text=wrap, vertical="top")
            c.border = BOX
            if ri % 2 == 0:
                c.fill = ALT_FILL
            if col_name == "Category" and val:
                cat_color = CAT_COLOR.get(val, "")
                if cat_color:
                    c.font = Font(color=cat_color, bold=True, size=10)

    widths = [12, 12, 12, 42, 42, 30, 24, 20, 20, 18, 14, 14, 14, 24, 18, 18, 24, 20, 20, 20, 12, 14, 14]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(cols))}2"


def write_summary_sheet(ws, rows: list[dict], month_label: str,
                        prev_counts: dict = None):
    ws.title = "Monthly Summary"
    counts = defaultdict(int)
    for row in rows:
        counts[row["Category"]] += 1

    sorted_cats = [(name, counts[name]) for name in ALL_CATS]
    total = sum(count for _, count in sorted_cats)

    ws.merge_cells("A1:F1")
    title = ws["A1"]
    title.value = f"NILE Pipeline Issue Summary — {month_label}"
    title.font = Font(bold=True, size=16, color="1F3864")
    title.alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 28

    ws["A2"].value = f"Total Tickets: {total}"
    ws["A2"].font = Font(bold=True, size=12, color="444444")
    ws.row_dimensions[2].height = 20

    headings = ["Category", "Count", "% Share", "vs Last Month", "Color Tag"]
    for ci, heading in enumerate(headings, 1):
        c = ws.cell(row=4, column=ci, value=heading)
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BOX
    ws.row_dimensions[4].height = 22

    for ri, (cat_name, count) in enumerate(sorted_cats, start=5):
        pct   = round(count / total * 100, 1) if total else 0
        color = CAT_COLOR.get(cat_name, "D9D9D9")
        alt   = ri % 2 == 0

        c1 = ws.cell(row=ri, column=1, value=cat_name)
        c1.alignment = Alignment(horizontal="left")
        c1.border = BOX
        if alt:
            c1.fill = ALT_FILL

        c2 = ws.cell(row=ri, column=2, value=count)
        c2.alignment = Alignment(horizontal="center")
        c2.font = Font(bold=True)
        c2.border = BOX
        if alt:
            c2.fill = ALT_FILL

        c3 = ws.cell(row=ri, column=3, value=f"{pct}%")
        c3.alignment = Alignment(horizontal="center")
        c3.border = BOX
        if alt:
            c3.fill = ALT_FILL

        # vs Last Month (column 4)
        prev     = (prev_counts or {}).get(cat_name, 0)
        diff     = count - prev
        if prev == 0 and count > 0:
            vs_txt   = "New"
            vs_color = "185FA5"
        elif diff > 0:
            vs_txt   = f"▲ +{diff}"
            vs_color = "3B6D11"
        elif diff < 0:
            vs_txt   = f"▼ {diff}"
            vs_color = "A32D2D"
        else:
            vs_txt   = "— same"
            vs_color = "888780"

        c4 = ws.cell(row=ri, column=4, value=vs_txt)
        c4.font      = Font(bold=True, size=11, color=vs_color)
        c4.alignment = Alignment(horizontal="center", vertical="center")
        c4.border    = BOX
        if alt:
            c4.fill = ALT_FILL

        # Color Tag moved to column 5
        color_cell = ws.cell(row=ri, column=5, value="")
        color_cell.fill   = PatternFill("solid", fgColor=color)
        color_cell.border = BOX

    total_row = 5 + len(sorted_cats)
    if total_row > 5:
        t1 = ws.cell(row=total_row, column=1, value="Monthly Total")
        t1.font   = Font(bold=True)
        t1.border = BOX

        t2 = ws.cell(row=total_row, column=2, value=total)
        t2.font      = Font(bold=True)
        t2.alignment = Alignment(horizontal="center")
        t2.border    = BOX

        t3 = ws.cell(row=total_row, column=3, value="100%")
        t3.alignment = Alignment(horizontal="center")
        t3.border    = BOX

        t4 = ws.cell(row=total_row, column=4, value="")
        t4.border = BOX

        t5 = ws.cell(row=total_row, column=5, value="")
        t5.border = BOX

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14

    chart_end = 4 + len(sorted_cats)
    if sorted_cats:
        bar = BarChart()
        bar.type = "col"
        bar.title = f"Issue Count by Category — {month_label}"
        bar.y_axis.title = "Count"
        bar.style = 10
        bar.width = 26
        bar.height = 15

        data = Reference(ws, min_col=2, min_row=4, max_row=chart_end)
        cats = Reference(ws, min_col=1, min_row=5, max_row=chart_end)
        bar.add_data(data, titles_from_data=True)
        bar.set_categories(cats)

        for idx, (cat_name, _) in enumerate(sorted_cats):
            pt = DataPoint(idx=idx)
            pt.graphicalProperties.solidFill = CAT_COLOR.get(cat_name, "D9D9D9")
            bar.series[0].dPt.append(pt)

        ws.add_chart(bar, "F4")

        pie = PieChart()
        pie.title = "Issue Distribution"
        pie.style = 10
        pie.width = 18
        pie.height = 14

        pie.add_data(data, titles_from_data=True)
        pie.set_categories(cats)
        for idx, (cat_name, _) in enumerate(sorted_cats):
            pt = DataPoint(idx=idx)
            pt.graphicalProperties.solidFill = CAT_COLOR.get(cat_name, "D9D9D9")
            pie.series[0].dPt.append(pt)

        ws.add_chart(pie, "F24")

    ws.freeze_panes = "A5"


def write_uncategorized_sheet(ws, rows: list[dict]):
    ws.title = "Needs Review"
    uncategorized = [r for r in rows if r["Category"] == OTHER_CATEGORY["name"]]

    if not uncategorized:
        ws["A1"].value = "All tickets were successfully categorized."
        ws["A1"].font = Font(bold=True, color="00B050")
        return

    ws.merge_cells("A1:G1")
    header = ws["A1"]
    header.value = f"Uncategorized Tickets — needs manual category ({len(uncategorized)} tickets)"
    header.font = Font(bold=True, size=13, color="C00000")
    header.alignment = Alignment(horizontal="center")

    cols = ["Key", "Summary", "Category", "Status", "Created", "Assignee", "Reviewer"]
    for ci, col_name in enumerate(cols, 1):
        c = ws.cell(row=2, column=ci, value=col_name)
        c.fill = PatternFill("solid", fgColor="C00000")
        c.font = Font(color="FFFFFF", bold=True)
        c.border = BOX

    for ri, row in enumerate(uncategorized, 3):
        for ci, col_name in enumerate(cols, 1):
            c = ws.cell(row=ri, column=ci, value=row.get(col_name, ""))
            c.border = BOX
            if ri % 2 == 0:
                c.fill = PatternFill("solid", fgColor="FFF2F2")
            c.alignment = Alignment(wrap_text=(col_name == "Summary"), vertical="top")

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 60
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 18
    ws.column_dimensions["G"].width = 16
