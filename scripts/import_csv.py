"""Start a plan from a list (B-914): an outline (.txt) or a CSV (this planner's columns or an MS Project export).

    python scripts/import_csv.py tasks.csv Project_Planner.xlsx [--project-type "IT infrastructure"] [--recalc]
    python scripts/import_csv.py outline.txt Project_Planner.xlsx

Outline: one task per line, indented (tab or spaces) or bulleted for levels, with optional "; 5d; after 3; Alex Chen".
The browser page docs/export.html does the same (choose the file, or paste the outline).
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from planner.build import Sizes, make_workbook, save  # noqa: E402
from planner.list_import import ListImportError, read  # noqa: E402
from planner.settings import PRESETS  # noqa: E402


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("source", help="a .csv or an outline .txt")
    p.add_argument("out", help="the workbook to write")
    p.add_argument("--project-type", choices=list(PRESETS), help="Start > Project type")
    p.add_argument("--recalc", action="store_true", help="recalculate with LibreOffice so the file carries values")
    a = p.parse_args(argv)
    try:
        project, warnings = read(Path(a.source).read_bytes(), a.source)
    except ListImportError as e:
        raise SystemExit(f"error: {e}")
    if a.project_type:
        project.settings["ProjectType"] = a.project_type
    sizes = Sizes(tasks=max(200, len(project.tasks) + 50), resources=max(50, len(project.resources) + 10))
    save(make_workbook(sizes=sizes, data=project), a.out, a.recalc)
    print(f"{len(project.tasks)} tasks, {len(project.resources)} resources")
    for w in warnings:
        print("note:", w)
    print("saved", a.out)


if __name__ == "__main__":
    main()
