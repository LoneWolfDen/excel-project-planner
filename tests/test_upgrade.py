"""Moving a plan into the latest template (planner/upgrade.py, scripts/upgrade_planner.py).

The fixtures are the sample workbooks as committed at v1.4 (two predecessors, "Start No Earlier Than", no
Assignments sheet) and v1.6, with the values Excel / LibreOffice calculated then. Upgraded and recalculated, the
plan must schedule exactly as it did, and everything typed must come across.
"""
import datetime as dt
import io
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))
from planner import upgrade  # noqa: E402
from test_planner import SOFFICE, build, errors_in, recalc  # noqa: E402
from test_mspdi import table_rows  # noqa: E402

FIXTURES = ROOT / "tests" / "fixtures"
SCHEDULE = ["Task Name", "Level", "Duration", "Forecast Start", "Forecast Finish", "Planned Start", "Late Finish",
            "Total Slack", "Work", "Cost", "Status", "Critical", "Progress"]
LOGS = [("RAID", "RAIDLog", "Title"), ("Decisions", "DecisionLog", "Decision"), ("Changes", "ChangeLog", "Title"),
        ("History", "StatusHistory", "Status Date")]


def log(wb, sheet, table, column):
    return [r for r in table_rows(wb, sheet, table) if r[column]]


class ReadTest(unittest.TestCase):
    def test_reads_an_older_layout_by_header(self):
        project, copies, sizes, warnings, version = upgrade.read_planner((FIXTURES / "planner_v1.4.xlsx").read_bytes())
        self.assertEqual(version, "1.6 or earlier")                 # no PlannerVersion before v1.7
        self.assertEqual(len(project.tasks), 30)
        uat = project.tasks[21]
        self.assertEqual((uat["Name"], uat["Cons"]), ("User acceptance testing", dt.datetime(2026, 12, 7)))
        self.assertEqual((project.tasks[8]["P1"], project.tasks[8]["P2"]), (7, 8))
        self.assertEqual(project.settings["ProjectName"], "Customer Portal Implementation")
        self.assertNotIn("ProjectFinish", project.settings)             # calculated settings stay the template's
        self.assertEqual(project.resources[0][:2], ("Priya Sharma", "PS"))
        self.assertIn("Casey Kim", project.leave)
        self.assertEqual(project.assignments, [])                       # v1.4 had no Assignments sheet
        self.assertIn("RAIDLog", copies)
        self.assertEqual(warnings, [])

    def test_current_workbooks_report_their_version(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "new.xlsx"
            build(path)
            self.assertEqual(upgrade.read_planner(path.read_bytes())[4], "1.7")

    def test_not_a_planner(self):
        with self.assertRaisesRegex(upgrade.UpgradeError, "isn't an Excel workbook"):
            upgrade.read_planner(b"not a zip")
        out = io.BytesIO()
        openpyxl.Workbook().save(out)
        with self.assertRaisesRegex(upgrade.UpgradeError, "no Tasks table"):
            upgrade.read_planner(out.getvalue())

    def test_custom_dropdown_values_are_kept_and_new_ones_added(self):
        wb = openpyxl.load_workbook(FIXTURES / "planner_v1.4.xlsx")
        ws = wb["Lists"]
        table = ws.tables["ListRAIDType"]
        last = openpyxl.utils.cell.range_boundaries(table.ref)
        ws.cell(last[3], last[0]).value = "Escalation"                  # the user replaced the last RAID type
        old = io.BytesIO()
        wb.save(old)
        new = openpyxl.load_workbook(io.BytesIO(upgrade.convert(old.getvalue())[0]))
        values = [r["RAID Type"] for r in table_rows(new, "Lists", "ListRAIDType") if r["RAID Type"]]
        self.assertEqual(values, ["Risk", "Assumption", "Issue", "Dependency", "Escalation"])
        periods = [r["Gantt Days per Column"] for r in table_rows(new, "Lists", "ListPeriod")]
        self.assertEqual(periods, [1, 7, 30, 91])                       # v1.6 values the v1.4 list didn't have


@unittest.skipUnless(SOFFICE, "LibreOffice not installed")
class UpgradeRecalcTest(unittest.TestCase):
    """Upgrade, recalculate, compare with what the old workbook calculated."""

    def check(self, fixture):
        old = openpyxl.load_workbook(FIXTURES / fixture, data_only=True)
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "Project_Planner.xlsx"
            result = subprocess.run([sys.executable, str(ROOT / "scripts" / "upgrade_planner.py"),
                                     str(FIXTURES / fixture), str(path)], capture_output=True, text=True, check=True)
            self.assertIn("saved", result.stdout)
            new = recalc(path, tmp, status_date=False)                  # the old workbook's own status date
        self.assertEqual(errors_in(new), [])
        before = {r["UID"]: r for r in table_rows(old, "Tasks", "Tasks") if r["Task Name"]}
        after = {r["UID"]: r for r in table_rows(new, "Tasks", "Tasks") if r["Task Name"]}
        self.assertEqual(sorted(after), sorted(before))
        diffs = [(u, k, before[u][k], after[u][k]) for u in before for k in SCHEDULE if before[u][k] != after[u][k]]
        self.assertEqual(diffs, [])
        for sheet, table, column in LOGS:
            self.assertEqual(log(new, sheet, table, column), log(old, sheet, table, column), table)
        leave = lambda wb: [(r["Resource"], r["Nov-26"]) for r in table_rows(wb, "Leave", "LeaveTable") if r["Resource"]]
        self.assertEqual(leave(new), leave(old))
        rag = lambda wb: [(r["Area"], r["PM RAG"], r["Comment"]) for r in table_rows(wb, "Status Report", "StatusRAG")]
        self.assertEqual(rag(new), rag(old))
        goals = lambda wb: [r["Goal"] for r in table_rows(wb, "Sprints", "SprintTable")]
        self.assertEqual(goals(new), goals(old))
        return new

    def test_from_v1_4(self):
        self.check("planner_v1.4.xlsx")

    def test_from_v1_6(self):
        new = self.check("planner_v1.6.xlsx")
        extra = [(r["Task UID"], r["Resource"]) for r in table_rows(new, "Assignments", "AssignmentTable") if r["Task UID"]]
        self.assertEqual(extra, [(7, "Jordan Lee"), (17, "Casey Kim"), (21, "Sam Rivera"), (22, "Taylor Morgan")])
        self.assertEqual(new.sheetnames.count("MSP_Tasks"), 0)


if __name__ == "__main__":
    unittest.main()
