"""
run.py — NILE Pipeline Issue Tracker launcher.

Usage:
  python run.py                        # auto-picks latest file in input/
  python run.py input/myfile.csv       # use a specific file
  python run.py input/Service_Desk.html
"""
from nile_pipeline_tracker.main import main

if __name__ == "__main__":
    main()
