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
from planner import settings, upgrade  # noqa: E402
from test_planner import SOFFICE, build, errors_in, recalc  # noqa: E402
from test_mspdi import table_rows  # noqa: E402

FIXTURES = ROOT / "tests" / "fixtures"
SCHEDULE = ["Task Name", "Level", "Duration", "Start", "Finish", "Planned Start", "Late Finish",
            "Total Slack", "Work", "Cost", "Status", "Critical", "Progress"]
OLD_LOGS = {"RAIDLog": "RAID", "DecisionLog": "Decisions", "ChangeLog": "Changes"}      # sheets before v1.7


def log(wb, sheet, table, column):
    return [r for r in table_rows(wb, sheet, table) if r[column]]


def old_log_items(old):
    """The v1.6 logs' rows under the Log's headers, as the upgrade should have written them (Ref left out: decisions
    and changes are renumbered), with what the old workbook calculated for them (score, rating, linked task)."""
    items = []
    for table, (kind, headers) in upgrade.LOGS.items():
        for r in table_rows(old, OLD_LOGS[table], table):
            r = {headers.get(h, h): v for h, v in r.items() if not h.startswith("_")}
            if r.get("Title"):
                items.append({"Type": kind, **r} if kind else r)
    return items


class ReadTest(unittest.TestCase):
    def test_reads_an_older_layout_by_header(self):
        project, copies, sizes, warnings, version = upgrade.read_planner((FIXTURES / "planner_v1.4.xlsx").read_bytes())
        self.assertEqual(version, "1.6 or earlier")                 # no PlannerVersion before v1.7
        self.assertEqual(len(project.tasks), 30)
        uat = project.tasks[21]
        self.assertEqual((uat["Name"], uat["Cons"]), ("User acceptance testing", dt.datetime(2026, 12, 7)))
        self.assertEqual(project.tasks[8]["PredIn"], "7, 8")                    # Pred1 / Pred2 UID -> one cell
        self.assertEqual((project.tasks[6]["PredIn"], project.tasks[17]["PredIn"]), ("6SS+3d", "17FF+1d"))
        self.assertEqual(project.settings["ProjectName"], "Customer Portal Implementation")
        self.assertNotIn("ProjectFinish", project.settings)             # calculated settings stay the template's
        self.assertEqual(project.resources[0][:2], ("Priya Sharma", "PS"))
        self.assertIn("Casey Kim", project.leave)
        self.assertEqual(project.assignments, [])                       # v1.4 had no Assignments sheet
        self.assertIn("Log", copies)                                     # RAID, decision and change logs, merged
        self.assertEqual([w.split(" refs")[0] for w in warnings], ["Decision", "Change"])   # renumbered, and said so
        self.assertIn("Decision 2 is now Decision 8", warnings[0])

    def test_current_workbooks_report_their_version(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "new.xlsx"
            build(path)
            self.assertEqual(upgrade.read_planner(path.read_bytes())[4], "1.7")

    def test_a_custom_week_becomes_custom(self):
        wb = openpyxl.load_workbook(FIXTURES / "planner_v1.4.xlsx")
        sheet, ref = wb.defined_names["WorkMon"].attr_text.split("!")
        wb[sheet][ref.replace("$", "")].value = "No"                     # Tue-Fri: no preset matches
        old = io.BytesIO()
        wb.save(old)
        self.assertEqual(upgrade.read_planner(old.getvalue())[0].settings["WorkingWeek"], "Custom")
        fixture = upgrade.read_planner((FIXTURES / "planner_v1.4.xlsx").read_bytes())[0]
        self.assertEqual(fixture.settings["WorkingWeek"], "Mon-Fri")
        self.assertEqual(settings.week_preset({"WorkSat": "Yes"}), "Mon-Sat")

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
        table = ws.tables["ListRAIDType"]                               # now ListLogType
        last = openpyxl.utils.cell.range_boundaries(table.ref)
        ws.cell(last[3], last[0]).value = "Escalation"                  # the user replaced the last RAID type
        old = io.BytesIO()
        wb.save(old)
        new = openpyxl.load_workbook(io.BytesIO(upgrade.convert(old.getvalue())[0]))
        values = [r["Log Type"] for r in table_rows(new, "Settings", "ListLogType") if r["Log Type"]]
        self.assertEqual(values, ["Risk", "Assumption", "Issue", "Dependency", "Decision", "Change", "Escalation"])
        periods = [r["Gantt Days per Column"] for r in table_rows(new, "Settings", "ListPeriod")]
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
        was = {"Start": "Forecast Start", "Finish": "Forecast Finish"}             # their names before v1.7
        diffs = [(u, k, before[u][was.get(k, k)], after[u][k]) for u in before for k in SCHEDULE
                 if before[u][was.get(k, k)] != after[u][k]]
        self.assertEqual(diffs, [])
        self.assertEqual(log(new, "History", "StatusHistory", "Status Date"),
                         log(old, "History", "StatusHistory", "Status Date"))
        items = old_log_items(old)
        logged = log(new, "Log", "Log", "Title")
        self.assertEqual(len(logged), len(items))
        for want, got in zip(items, logged):
            self.assertEqual({k: got[k] for k in want if k != "Ref"}, {k: v for k, v in want.items() if k != "Ref"})
        raid = [(r["Ref"], r["Type"]) for r in log(old, "RAID", "RAIDLog", "Title")]
        self.assertEqual([(r["Ref"], r["Type"]) for r in logged[:len(raid)]], raid)     # RAID items keep their Ref
        self.assertEqual([r["Ref"] for r in logged], list(range(1, len(logged) + 1)))
        before = {r["Resource"]: r["Nov-26"] for r in table_rows(old, "Leave", "LeaveTable") if r["Resource"]}
        after = {r["Resource Name"]: r["Nov-26"] for r in table_rows(new, "Resources", "ResourceTable")
                 if r["Resource Name"] in before}                       # the Leave sheet's rows, now on Resources
        self.assertEqual(after, before)
        rag = lambda wb, sheet: [(r["Area"], r["PM RAG"], r["Comment"]) for r in table_rows(wb, sheet, "StatusRAG")]
        self.assertEqual(rag(new, "Status"), rag(old, "Status Report"))
        goals = lambda wb, sheet: [r["Goal"] for r in table_rows(wb, sheet, "SprintTable")]
        self.assertEqual(goals(new, "Agile"), goals(old, "Sprints"))
        return new

    def test_from_v1_4(self):
        self.check("planner_v1.4.xlsx")

    def test_from_v1_6(self):
        new = self.check("planner_v1.6.xlsx")
        people = {r["UID"]: r["Resources"] for r in table_rows(new, "Tasks", "Tasks") if r["UID"] in (7, 17, 21, 22)}
        self.assertEqual(people, {7: "Alex Chen, Jordan Lee[50%]", 17: "Sam Rivera, Casey Kim[50%]",   # Assignments rows
                                  21: "Taylor Morgan, Sam Rivera[50%]", 22: "Morgan Patel, Taylor Morgan"})
        self.assertEqual(new.sheetnames.count("MSP_Tasks"), 0)


if __name__ == "__main__":
    unittest.main()
