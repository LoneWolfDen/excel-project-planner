"""Starting a plan from a list (planner/list_import.py, scripts/import_csv.py): B-914."""
import datetime as dt
import io
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))
from planner import list_import  # noqa: E402
from test_mspdi import table_rows  # noqa: E402
from test_planner import SOFFICE, d, errors_in, recalc  # noqa: E402

OUTLINE = """Design
  Requirements; 5d; Priya Sharma
  Wireframes; 3; after 2; Casey Kim
  Sign-off; milestone; after 3
Build
  - API; 2 wks; after 4; Jordan Lee, SR[50%]
  - Screens; 8d; after 6SS+2d; Sam Rivera
"""
PLANNER_CSV = """Level,Task Name,Duration,Predecessors,Resources,Notes
1,Discovery,,,,
2,Interviews,4,,Morgan Patel,Ten users
2,Synthesis,2,2,Morgan Patel,
1,Decision,0,3,,
"""
MSP_CSV = """ID,Unique ID,Name,Duration,Start,Finish,Predecessors,Resource Names,Outline Level,% Complete
1,10,Plan,6 days,Mon 07/09/26,Mon 14/09/26,,,1,0%
2,11,Scope,3 days,Mon 07/09/26,Wed 09/09/26,,Priya Sharma,2,100%
3,12,Kick-off,0 days,Wed 09/09/26,Wed 09/09/26,2,,2,0%
4,13,Design,5 days,Thu 10/09/26,Wed 16/09/26,3FS+1 day,Alex Chen;Casey Kim[50%],2,0%
"""


class ParseTest(unittest.TestCase):
    def test_durations_and_links(self):
        self.assertEqual([list_import.days(t) for t in ("5", "5d", "5 days", "2 wks", "16 hrs", "0 days", "1 mo")],
                         [5, 5, 5, 10, 2, 0, 20])
        self.assertIsNone(list_import.days("soon"))
        self.assertEqual(list_import.links("3, 5SS+2d, 4FF-1"), ("3, 5SS+2d, 4FF-1d", []))
        self.assertEqual(list_import.links("3FS+2 days;5", lambda n: n + 10), ("13FS+2d, 15", []))
        self.assertEqual(list_import.links("3, later"), ("3", ["later"]))

    def test_outline(self):
        rows, warnings = list_import.from_outline(OUTLINE)
        self.assertEqual(warnings, [])
        self.assertEqual([(r["UID"], r["Level"], r["Name"]) for r in rows],
                         [(1, 1, "Design"), (2, 2, "Requirements"), (3, 2, "Wireframes"), (4, 2, "Sign-off"),
                          (5, 1, "Build"), (6, 2, "API"), (7, 2, "Screens")])
        self.assertEqual((rows[3]["Dur"], rows[3]["PredIn"]), (0, "3"))
        self.assertEqual((rows[5]["Dur"], rows[5]["ResIn"]), (10, "Jordan Lee, SR[50%]"))
        self.assertEqual(rows[6]["PredIn"], "6SS+2d")
        numbered, _ = list_import.from_outline("1 Plan\n1.1 Scope; 2d\n1.2 Approve; 0d; after 2\n2 Do")
        self.assertEqual([r["Level"] for r in numbered], [1, 2, 2, 1])

    def test_ms_project_csv(self):
        rows, warnings, start = list_import.from_csv(MSP_CSV)
        self.assertEqual((warnings, start), ([], dt.date(2026, 9, 7)))
        self.assertEqual([r["UID"] for r in rows], [10, 11, 12, 13])           # Unique ID kept
        self.assertNotIn("Dur", rows[0])                                       # a summary rolls up
        self.assertEqual((rows[1]["Dur"], rows[1]["Pct"]), (3, 1.0))
        self.assertEqual(rows[3]["PredIn"], "12FS+1d")                         # by ID 3 -> UID 12
        self.assertEqual(rows[3]["ResIn"], "Alex Chen, Casey Kim[50%]")

    def test_errors(self):
        with self.assertRaisesRegex(list_import.ListImportError, "no tasks"):
            list_import.read("   \n\n")
        with self.assertRaisesRegex(list_import.ListImportError, "Task Name"):
            list_import.from_csv("a,b,c\n1,2,3\n")


class WorkbookTest(unittest.TestCase):
    def build(self, text, name="list"):
        xlsx, warnings = list_import.convert(text, name)
        return openpyxl.load_workbook(io.BytesIO(xlsx)), warnings

    def test_planner_csv(self):
        wb, warnings = self.build(PLANNER_CSV.encode(), "tasks.csv")
        self.assertEqual(warnings, [])
        tasks = [r for r in table_rows(wb, "Tasks", "Tasks") if r["Task Name"]]
        self.assertEqual([(r["Level"], r["Task Name"], r["Duration"], r["Predecessors"]) for r in tasks],
                         [(1, "Discovery", None, None), (2, "Interviews", 4, None), (2, "Synthesis", 2, "2"),
                          (1, "Decision", 0, "3")])
        people = [r["Resource Name"] for r in table_rows(wb, "Resources", "ResourceTable") if r["Resource Name"]]
        self.assertEqual(people, ["Morgan Patel"])

    def test_links_below_are_reported(self):
        _, warnings = self.build("A; 2d; after 2\nB; 3d")
        self.assertIn("1 links point to a task listed below its successor", " ".join(warnings))

    @unittest.skipUnless(SOFFICE, "LibreOffice not installed")
    def test_outline_schedules(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "outline.xlsx"
            path.write_bytes(list_import.convert(OUTLINE)[0])
            wb = recalc(path, tmp, status_date=False)
            self.assertEqual(errors_in(wb), [])
            t = {r["UID"]: r for r in table_rows(wb, "Tasks", "Tasks") if r["Task Name"]}
            self.assertEqual(d(t[3]["Start"]), d(t[2]["Finish"]) + dt.timedelta(days=1 if
                                                                               d(t[2]["Finish"]).weekday() < 4 else 3))
            self.assertEqual(t[6]["Resource Names"], "Jordan Lee, Sam Rivera[50%]")       # initials resolved
            self.assertFalse([r["Issues"] for r in t.values() if "not on Resources" in str(r["Issues"])])


if __name__ == "__main__":
    unittest.main()
