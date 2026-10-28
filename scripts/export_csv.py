"""Export the plan as CSV for Jira, Azure DevOps or Microsoft Planner (B-706).

    python scripts/export_csv.py Project_Planner.xlsx tasks.csv --format jira      # Jira: System > External System Import > CSV
    python scripts/export_csv.py Project_Planner.xlsx tasks.csv --format azure     # Azure DevOps: Boards > Import Work Items
    python scripts/export_csv.py Project_Planner.xlsx tasks.csv --format planner   # Planner / Project for the web

Reads the workbook's calculated values: open and save it in Excel first (or build it with --recalc).
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from planner.exports import FORMATS, to_csv  # noqa: E402
from planner.mspdi import ExportError  # noqa: E402


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("workbook")
    p.add_argument("out")
    p.add_argument("--format", choices=list(FORMATS), default="jira")
    a = p.parse_args(argv)
    try:
        Path(a.out).write_text(to_csv(Path(a.workbook).read_bytes(), a.format), newline="")
    except ExportError as e:
        raise SystemExit(str(e))
    print("saved", a.out)


if __name__ == "__main__":
    main()
