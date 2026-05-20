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


def _cell(ws, row: int, col: int, value="", bold=False, center=False, fill=None, font=None, wrap=False):
    c = ws.cell(row=row, column=col, value=value)
    if bold:
        c.font = Font(bold=True)
    if center:
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
    elif wrap:
        c.alignment = Alignment(wrap_text=True)
    if fill:
        c.fill = fill
    if font:
        c.font = font
    c.border = BOX
    return c


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

    widths = [12, 12, 12, 42, 42, 30, 24, 20, 20, 18, 14, 14, 14, 24, 18, 18, 24, 20, 20, 20, 12, 14, 14]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(cols))}2"


def write_summary_sheet(ws, rows: list[dict], month_label: str):
    ws.title = "Monthly Summary"
    counts = defaultdict(int)
    for row in rows:
        counts[row["Category"]] += 1

    sorted_cats = [(name, counts[name]) for name in ALL_CATS if counts[name] > 0]
    total = sum(count for _, count in sorted_cats)

    ws.merge_cells("A1:E1")
    title = ws["A1"]
    title.value = f"NILE Pipeline Issue Summary — {month_label}"
    title.font = Font(bold=True, size=16, color="1F3864")
    title.alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 28

    ws["A2"].value = f"Total Tickets: {total}"
    ws["A2"].font = Font(bold=True, size=12, color="444444")
    ws.row_dimensions[2].height = 20

    headings = ["Category", "Count", "% Share", "Color Tag"]
    for ci, heading in enumerate(headings, 1):
        c = ws.cell(row=4, column=ci, value=heading)
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BOX
    ws.row_dimensions[4].height = 22

    for ri, (cat_name, count) in enumerate(sorted_cats, start=5):
        pct = round(count / total * 100, 1) if total else 0
        color = CAT_COLOR.get(cat_name, "D9D9D9")
        alt = ri % 2 == 0

        ws.cell(row=ri, column=1, value=cat_name).alignment = Alignment(horizontal="left")
        ws.cell(row=ri, column=1).border = BOX
        if alt:
            ws.cell(row=ri, column=1).fill = ALT_FILL

        ws.cell(row=ri, column=2, value=count).alignment = Alignment(horizontal="center")
        ws.cell(row=ri, column=2).font = Font(bold=True)
        ws.cell(row=ri, column=2).border = BOX
        if alt:
            ws.cell(row=ri, column=2).fill = ALT_FILL

        ws.cell(row=ri, column=3, value=f"{pct}%").alignment = Alignment(horizontal="center")
        ws.cell(row=ri, column=3).border = BOX
        if alt:
            ws.cell(row=ri, column=3).fill = ALT_FILL

        color_cell = ws.cell(row=ri, column=4, value="")
        color_cell.fill = PatternFill("solid", fgColor=color)
        color_cell.border = BOX

    total_row = 5 + len(sorted_cats)
    if total_row > 5:
        ws.cell(row=total_row, column=1, value="Monthly Total")
        ws.cell(row=total_row, column=1).font = Font(bold=True)
        ws.cell(row=total_row, column=1).border = BOX
        ws.cell(row=total_row, column=2, value=total).font = Font(bold=True)
        ws.cell(row=total_row, column=2).alignment = Alignment(horizontal="center")
        ws.cell(row=total_row, column=2).border = BOX
        ws.cell(row=total_row, column=3, value="100%").alignment = Alignment(horizontal="center")
        ws.cell(row=total_row, column=3).border = BOX
        ws.cell(row=total_row, column=4, value="")
        ws.cell(row=total_row, column=4).border = BOX

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 14

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
