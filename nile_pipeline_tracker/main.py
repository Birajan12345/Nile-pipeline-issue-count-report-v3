import csv
import os
import sys
from collections import Counter

import openpyxl

from .config import CSV_PATH, OUTPUT_DIR, REPORT_MONTH_LABEL
from .csv_reader import detect_month_label, parse_rows, read_csv, read_html
from .excel_writer import write_raw_sheet, write_summary_sheet, write_uncategorized_sheet


def read_source(input_path: str):
    _, ext = os.path.splitext(input_path.lower())
    if ext in (".html", ".htm"):
        return read_html(input_path)
    return read_csv(input_path)


def main(argv=None):
    argv = argv or sys.argv
    source_path = argv[1] if len(argv) > 1 else CSV_PATH

    print(f"\nNILE Pipeline Issue Tracker")
    print(f"{'─' * 40}")
    print(f"  Reading file : {source_path}")

    raw_rows, col_map = read_source(source_path)
    rows = parse_rows(raw_rows, col_map)

    month_label = REPORT_MONTH_LABEL or detect_month_label(rows)
    print(f"  Month label : {month_label}")

    cat_counts = Counter(r["Category"] for r in rows)
    print(f"\n  Category Breakdown:")
    from .config import ALL_CATS

    for name in ALL_CATS:
        if cat_counts[name]:
            print(f"    {name:<32} {cat_counts[name]}")
    print(f"    {'─' * 36}")
    print(f"    {'TOTAL':<32} {len(rows)}")

    safe_label = month_label.replace(" ", "_")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    wb = openpyxl.Workbook()
    write_raw_sheet(wb.active, rows, month_label)
    write_summary_sheet(wb.create_sheet(), rows, month_label)
    write_uncategorized_sheet(wb.create_sheet(), rows)

    xlsx_name = f"Nile_Pipeline_Issue_Count_{safe_label}.xlsx"
    xlsx_path = os.path.join(OUTPUT_DIR, xlsx_name)
    wb.save(xlsx_path)
    print(f"\n  ✓ Excel report saved: {xlsx_path}")

    csv_name = f"Nile_Pipeline_Issue_Count_{safe_label}.csv"
    csv_path = os.path.join(OUTPUT_DIR, csv_name)
    with open(csv_path, 'w', newline='', encoding='utf-8') as f:
        fieldnames = [
            "Key", "Issue id", "Parent id", "Summary", "Description", "Comment", "Close Notes",
            "Categories", "New Categories", "Assign", "Status", "Created", "Updated",
            "Custom field (Assignment Group)", "Assignee", "Reporter", "Resolution",
            "Custom field (Comments)", "Inward issue link (Cloners)", "Outward issue link (Cloners)",
            "Inward issue link (Relates)", "Category", "Review Required",
        ]
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)

    print(f"  ✓ CSV report saved: {csv_path}")
    print(f"{'─' * 40}\n")


if __name__ == "__main__":
    main()
