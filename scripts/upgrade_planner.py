"""Move a plan made with an older version of the template into the latest one (B-901).

    python scripts/upgrade_planner.py Old_Planner.xlsx New_Planner.xlsx [--recalc]

Everything typed into the old workbook (tasks, people, assignments, leave, holidays, settings, logs, status history,
sprint goals, PM RAG and notes) is carried over by header name; the formulas are the latest template's. The browser
page docs/export.html does the same ("Move to the latest template").
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from planner import upgrade  # noqa: E402
from planner.build import save  # noqa: E402


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("old", help="the older planner workbook (.xlsx)")
    p.add_argument("out", help="the new workbook to write")
    p.add_argument("--recalc", action="store_true", help="recalculate with LibreOffice so the file carries values")
    a = p.parse_args(argv)
    try:
        project, copies, sizes, warnings, version = upgrade.read_planner(Path(a.old).read_bytes())
    except upgrade.UpgradeError as e:
        raise SystemExit(str(e))
    from planner.build import make_workbook
    wb = make_workbook(sizes=sizes, data=project)
    upgrade.copy_inputs(wb, copies, warnings)
    save(wb, a.out, a.recalc)
    print(f"read version {version}")
    for w in warnings:
        print("note:", w)
    print("saved", a.out)


if __name__ == "__main__":
    main()
