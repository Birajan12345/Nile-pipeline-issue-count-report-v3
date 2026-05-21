"""
main.py
Orchestrates the full pipeline:
  1. Auto-detect latest source file in input/
  2. Parse rows  (HTML / CSV / XLSX)
  3. Categorize  (scored keyword matching)
  4. Load previous month counts for trend comparison
  5. Write Excel report — 4 sheets all from same data source
  6. Write CSV report to output/
  7. Write timestamped run log to output/
"""
import csv
import logging
import os
import sys
from collections import Counter
from datetime import datetime

import openpyxl

from .config import ALL_CATS, OUTPUT_DIR, REPORT_MONTH_LABEL
from .csv_reader import detect_month_label, parse_rows, read_csv, read_html, read_xlsx
from .excel_writer import (
    load_prev_month_counts,
    write_cover_sheet,
    write_raw_sheet,
    write_summary_sheet,
    write_uncategorized_sheet,
)
from .file_finder import find_latest_source


def _setup_logger(output_dir: str, label: str) -> logging.Logger:
    """Write DEBUG+ to a timestamped .log file. Console output via print()."""
    os.makedirs(output_dir, exist_ok=True)
    safe     = label.replace(" ", "_")
    ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = os.path.join(output_dir, f"run_{safe}_{ts}.log")

    logger = logging.getLogger("nile_tracker")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()

    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter(
        "%(asctime)s  %(levelname)-7s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    ))
    logger.addHandler(fh)

    logger.info(f"Log file      : {log_path}")
    return logger


def read_source(input_path: str):
    """Dispatch to the correct reader based on file extension."""
    _, ext = os.path.splitext(input_path.lower())
    if ext in (".html", ".htm"):
        return read_html(input_path)
    if ext in (".xlsx", ".xls"):
        return read_xlsx(input_path)
    return read_csv(input_path)


def main(argv=None):
    argv = argv or sys.argv

    print(f"\n{'═' * 52}")
    print(f"  NILE Pipeline Issue Tracker")
    print(f"{'═' * 52}")

    # 1. Resolve source file
    if len(argv) > 1:
        source_path = argv[1]
        print(f"\n  Source file   : {source_path}  (from argument)")
    else:
        source_path = find_latest_source()

    # 2. Parse rows
    raw_rows, _ = read_source(source_path)
    rows = parse_rows(raw_rows)

    # 3. Month label — auto-detect when config value is blank
    month_label = REPORT_MONTH_LABEL or detect_month_label(rows)
    print(f"  Report month  : {month_label}")

    logger = _setup_logger(OUTPUT_DIR, month_label)
    logger.info(f"Source        : {source_path}")
    logger.info(f"Report month  : {month_label}")
    logger.info(f"Total rows    : {len(rows)}")

    # 4. Category counts — single source of truth for all 4 sheets
    cat_counts = Counter(r["Category"] for r in rows)

    # 5. Previous month counts for trend comparison
    prev_counts = load_prev_month_counts(OUTPUT_DIR)
    has_prev    = bool(prev_counts)
    print(f"  Previous data : {'found — trend column active' if has_prev else 'not found — trend column will show New'}")
    logger.info(f"Prev month    : {'loaded' if has_prev else 'none found'}")

    # 6. Category breakdown with trend indicators
    print(f"\n  Category breakdown:")
    for name in ALL_CATS:
        if cat_counts[name]:
            diff  = cat_counts[name] - prev_counts.get(name, 0)
            trend = f"  ▲ +{diff}" if diff > 0 else (f"  ▼ {diff}" if diff < 0 else "")
            print(f"    {name:<36} {cat_counts[name]}{trend}")
            logger.info(f"  {name:<36} {cat_counts[name]}")
    print(f"    {'─' * 42}")
    print(f"    {'TOTAL':<36} {len(rows)}")
    logger.info(f"  TOTAL                                  {len(rows)}")

    # 7. Build Excel — all 4 sheets from the same data
    safe_label = month_label.replace(" ", "_")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    wb = openpyxl.Workbook()

    # Sheet order: Cover Sheet opens first
    write_cover_sheet(wb.active, rows, month_label, dict(cat_counts), prev_counts)
    write_raw_sheet(wb.create_sheet(), rows, month_label)
    write_summary_sheet(wb.create_sheet(), rows, month_label, prev_counts)
    write_uncategorized_sheet(wb.create_sheet(), rows)

    xlsx_name = f"Nile_Pipeline_Issue_Count_{safe_label}.xlsx"
    xlsx_path = os.path.join(OUTPUT_DIR, xlsx_name)
    wb.save(xlsx_path)
    logger.info(f"Excel saved   : {xlsx_path}")
    print(f"\n  ✓ Excel report  → {xlsx_path}")

    # 8. Write CSV
    csv_name  = f"Nile_Pipeline_Issue_Count_{safe_label}.csv"
    csv_path  = os.path.join(OUTPUT_DIR, csv_name)
    fieldnames = [
        "Key", "Issue id", "Parent id", "Summary", "Description",
        "Comment", "Close Notes", "Categories", "New Categories",
        "Assign", "Status", "Created", "Updated",
        "Custom field (Assignment Group)", "Assignee", "Reporter",
        "Resolution", "Custom field (Comments)",
        "Inward issue link (Cloners)", "Outward issue link (Cloners)",
        "Inward issue link (Relates)", "Category", "Review Required",
    ]
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    logger.info(f"CSV saved     : {csv_path}")
    print(f"  ✓ CSV report    → {csv_path}")
    print(f"\n{'═' * 52}\n")


if __name__ == "__main__":
    main()
