"""
file_finder.py
Scans the input/ folder and returns the path of the
most recently modified supported source file.
"""
import os
import sys
from datetime import datetime

from .config import INPUT_DIR, SUPPORTED_EXTENSIONS


def find_latest_source() -> str:
    """Return path of the newest supported file in INPUT_DIR."""
    if not os.path.isdir(INPUT_DIR):
        print(f"\n  ERROR: Input folder not found: '{INPUT_DIR}'")
        print("  Create the folder and drop your Jira export inside it.")
        sys.exit(1)

    candidates = []
    skipped    = []

    for fname in os.listdir(INPUT_DIR):
        fpath = os.path.join(INPUT_DIR, fname)
        if not os.path.isfile(fpath):
            continue
        ext = os.path.splitext(fname)[1].lower()
        if ext in SUPPORTED_EXTENSIONS:
            candidates.append((os.path.getmtime(fpath), fpath, fname))
        else:
            skipped.append(fname)

    if not candidates:
        print(f"\n  ERROR: No supported file found in '{INPUT_DIR}/'")
        print(f"  Supported types: {', '.join(SUPPORTED_EXTENSIONS)}")
        if skipped:
            print(f"  Files found but skipped: {', '.join(skipped)}")
        print("\n  Drop your Jira export into input/ and re-run.")
        sys.exit(1)

    candidates.sort(reverse=True)
    mtime, path, name = candidates[0]

    print(f"\n  Input folder  : {INPUT_DIR}/")
    if len(candidates) > 1:
        print(f"  Files found   : {len(candidates)}")
        for i, (mt, _, fn) in enumerate(candidates):
            marker = "→ selected" if i == 0 else "  skipped "
            ts = datetime.fromtimestamp(mt).strftime("%Y-%m-%d %H:%M")
            print(f"    [{marker}]  {fn}  [{ts}]")
        print(f"  Auto-selected : {name}  (most recently modified)")
    else:
        ts = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
        print(f"  Source file   : {name}  [{ts}]")

    return path
