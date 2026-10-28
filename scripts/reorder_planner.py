"""Reorder a plan's tasks so every predecessor is above its successor (B-1015), keeping the outline and the UIDs.

    python scripts/reorder_planner.py Project_Planner.xlsx Reordered.xlsx [--recalc]

Tasks move only inside their phase, as whole sub-trees; when the outline makes it impossible, nothing is written and
the message names the tasks. Everything else typed in the workbook is carried over, as by upgrade_planner.py. The
browser page docs/export.html does the same ("Reorder tasks ...").
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from planner import reorder, upgrade  # noqa: E402
from planner.build import make_workbook, save  # noqa: E402


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("workbook", help="the planner workbook (.xlsx)")
    p.add_argument("out", help="the reordered workbook to write")
    p.add_argument("--recalc", action="store_true", help="recalculate with LibreOffice so the file carries values")
    a = p.parse_args(argv)
    try:
        project, copies, sizes, warnings, _version = upgrade.read_planner(Path(a.workbook).read_bytes())
        project.tasks, moved = reorder.order(project.tasks)
    except (upgrade.UpgradeError, reorder.ReorderError) as e:
        raise SystemExit(str(e))
    wb = make_workbook(sizes=sizes, data=project)
    upgrade.copy_inputs(wb, copies, warnings)
    save(wb, a.out, a.recalc)
    print(f"moved {len(moved)} tasks" + (f": {', '.join(map(str, moved))}" if moved else " (already in order)"))
    for w in warnings:
        print("note:", w)
    print("saved", a.out)


if __name__ == "__main__":
    main()
