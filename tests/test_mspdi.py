"""Tests for the MS Project XML (MSPDI) export: planner/mspdi.py, scripts/export_mspdi.py and docs/export.html.

The export reads the values saved in the workbook, so most tests run on the committed Project_Planner.xlsx and need
no LibreOffice. The browser test runs docs/export.html in headless Chromium when Node's playwright and a local copy of
Pyodide are available (PYODIDE_DIR; CI downloads it), and checks it writes exactly what the command line writes.
"""
import datetime as dt
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
from planner import mspdi  # noqa: E402
import build_export_page  # noqa: E402
from test_planner import BASELINE, FIRST, SOFFICE, SOURCE, build, columns, recalc, set_input  # noqa: E402

NS = "{http://schemas.microsoft.com/project}"
SAMPLE = ROOT / "Project_Planner.xlsx"


def parse(xml):
    return ET.fromstring(xml.encode("utf-8"))


def tasks_by_uid(root):
    return {int(t.find(NS + "UID").text): t for t in root.iter(NS + "Task")}


def text(element, tag):
    found = element.find(NS + tag)
    return None if found is None else found.text


def day(value):
    return dt.date.fromisoformat(value[:10])


class ExportTest(unittest.TestCase):
    """The committed sample workbook, exported and checked against its own Tasks sheet."""

    @classmethod
    def setUpClass(cls):
        cls.xml = mspdi.convert(SAMPLE.read_bytes())
        cls.root = parse(cls.xml)
        cls.tasks = tasks_by_uid(cls.root)
        wb = openpyxl.load_workbook(SAMPLE, data_only=True)
        ws = wb["Tasks"]
        col = columns(ws)
        cls.sheet = {}
        for r in range(FIRST, ws.max_row + 1):
            if ws.cell(r, col["Task Name"]).value:
                cls.sheet[ws.cell(r, col["UID"]).value] = {h: ws.cell(r, c).value for h, c in col.items()}

    def test_project_and_calendar(self):
        p = self.root
        self.assertEqual(p.tag, NS + "Project")
        self.assertEqual((text(p, "Title"), text(p, "Manager")), ("Customer Portal Implementation", "Priya Sharma"))
        self.assertEqual(text(p, "StartDate"), "2026-09-07T08:00:00")
        self.assertEqual((text(p, "MinutesPerDay"), text(p, "MinutesPerWeek")), ("480", "2400"))
        days = p.findall(f"{NS}Calendars/{NS}Calendar/{NS}WeekDays/{NS}WeekDay")
        week = {int(text(d, "DayType")): text(d, "DayWorking") for d in days if text(d, "DayType") != "0"}
        self.assertEqual(week, {1: "0", 2: "1", 3: "1", 4: "1", 5: "1", 6: "1", 7: "0"})     # Sunday = 1
        holidays = [day(d.find(f"{NS}TimePeriod/{NS}FromDate").text) for d in days if text(d, "DayType") == "0"]
        self.assertEqual(holidays, [dt.date(2026, 12, 24), dt.date(2026, 12, 25), dt.date(2026, 12, 31),
                                    dt.date(2027, 1, 1)])
        monday = next(d for d in days if text(d, "DayType") == "2")
        shifts = [(text(w, "FromTime"), text(w, "ToTime")) for w in monday.iter(NS + "WorkingTime")]
        self.assertEqual(shifts, [("08:00:00", "12:00:00"), ("13:00:00", "17:00:00")])

    def test_every_task_matches_the_tasks_sheet(self):
        self.assertEqual(sorted(self.tasks), sorted(self.sheet))
        for uid, t in self.tasks.items():
            row = self.sheet[uid]
            self.assertEqual(text(t, "Name"), row["Task Name"], uid)
            self.assertEqual(int(text(t, "ID")), row["ID"], uid)
            self.assertEqual(text(t, "WBS"), str(row["WBS"]), uid)
            self.assertEqual(int(text(t, "OutlineLevel")), row["Level"], uid)
            self.assertEqual(day(text(t, "Start")), row["Forecast Start"].date(), uid)
            self.assertEqual(day(text(t, "Finish")), row["Forecast Finish"].date(), uid)
            self.assertEqual(text(t, "Summary"), "1" if row["_IsSummary"] else "0", uid)
            self.assertEqual(int(text(t, "PercentComplete")), round(row["Progress"] * 100), uid)

    def test_durations_links_and_constraints(self):
        t = self.tasks
        self.assertEqual(text(t[6], "Duration"), "PT64H0M0S")                    # 8 days x 8 hours
        link = lambda uid: [(int(text(l, "PredecessorUID")), int(text(l, "Type")), int(text(l, "LinkLag")))
                            for l in t[uid].findall(NS + "PredecessorLink")]
        self.assertEqual(link(7), [(6, 3, 3 * 8 * 600)])                        # SS + 3 days, in tenths of a minute
        self.assertEqual(link(18), [(17, 0, 1 * 8 * 600)])                      # FF + 1 day
        self.assertEqual(link(9), [(7, 1, 0), (8, 1, 0)])                       # two FS links
        self.assertEqual(link(1), [])                                           # summaries carry no links
        self.assertEqual((text(t[22], "ConstraintType"), text(t[22], "ConstraintDate")), ("4", "2026-12-07T08:00:00"))
        self.assertEqual((text(t[28], "Manual"), text(t[28], "ConstraintType")), ("1", "0"))
        self.assertEqual(day(text(t[28], "Start")), dt.date(2026, 11, 16))
        self.assertEqual((text(t[4], "Milestone"), text(t[4], "Duration")), ("1", "PT0H0M0S"))
        self.assertEqual(text(t[19], "Deadline"), "2026-11-06T17:00:00")
        self.assertEqual((text(t[2], "ActualStart"), text(t[2], "ActualFinish"), text(t[2], "ActualCost")),
                         ("2026-09-07T08:00:00", "2026-09-10T17:00:00", "3040"))
        self.assertIsNone(t[2].find(NS + "Baseline"))                             # no baseline in the sample

    def test_custom_fields(self):
        values = lambda uid: {text(a, "FieldID"): text(a, "Value") for a in self.tasks[uid].findall(NS + "ExtendedAttribute")}
        self.assertEqual(values(17), {mspdi.TEXT1: "Sprint 4", mspdi.NUMBER1: "8"})
        self.assertEqual(values(4), {mspdi.TEXT1: "Sprint 1"})                     # a milestone has no points
        aliases = {text(a, "FieldID"): text(a, "Alias") for a in self.root.find(NS + "ExtendedAttributes")}
        self.assertEqual(aliases, {mspdi.TEXT1: "Sprint", mspdi.NUMBER1: "Story Points"})

    def test_resources_and_assignments(self):
        resources = {text(r, "Name"): r for r in self.root.iter(NS + "Resource")}
        self.assertEqual(len(resources), 7)
        self.assertEqual((text(resources["Priya Sharma"], "MaxUnits"), text(resources["Priya Sharma"], "StandardRate")),
                         ("0.5", "95"))
        uid_of = {text(r, "UID"): n for n, r in resources.items()}
        assignments = {int(text(a, "TaskUID")): (uid_of[text(a, "ResourceUID")], float(text(a, "Units")), text(a, "Work"))
                       for a in self.root.iter(NS + "Assignment")}
        expected = {u: r for u, r in self.sheet.items() if r["Resource"] and not r["_IsSummary"]}
        self.assertEqual(sorted(assignments), sorted(expected))
        self.assertEqual(assignments[18], ("Casey Kim", 0.5, "PT12H0M0S"))     # 3 days at 50%
        self.assertEqual(assignments[2], ("Priya Sharma", 1.0, "PT24H0M0S"))

    def test_elements_in_ms_project_order(self):
        # MS Project reports a file with elements out of order as corrupt; every tag must be in ORDER
        for element in self.root.iter():
            tag = element.tag.replace(NS, "")
            if tag in mspdi.ORDER:
                children = [c.tag.replace(NS, "") for c in element]
                unknown = [c for c in children if c not in mspdi.ORDER[tag]]
                self.assertEqual(unknown, [], tag)
                positions = [mspdi.ORDER[tag].index(c) for c in children]
                self.assertEqual(positions, sorted(positions), tag)

    def test_command_line(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "project.xml"
            subprocess.run([sys.executable, str(ROOT / "scripts" / "export_mspdi.py"), str(SAMPLE), str(out)],
                           check=True, capture_output=True)
            self.assertEqual(out.read_text(encoding="utf-8"), self.xml)


class ExportErrorTest(unittest.TestCase):
    def test_not_a_workbook(self):
        with self.assertRaisesRegex(mspdi.ExportError, "isn't an .xlsx"):
            mspdi.convert(b"not a zip file")

    def test_blank_template_needs_a_start_date(self):
        with self.assertRaisesRegex(mspdi.ExportError, "Project start date is empty"):
            mspdi.convert((ROOT / "Project_Planner_Blank.xlsx").read_bytes())

    def test_workbook_never_calculated(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "raw.xlsx"
            build(path)                                    # openpyxl writes formulas without values
            with self.assertRaisesRegex(mspdi.ExportError, "no calculated values"):
                mspdi.convert(path.read_bytes())

    def test_command_line_reports_the_problem(self):
        result = subprocess.run([sys.executable, str(ROOT / "scripts" / "export_mspdi.py"),
                                 str(ROOT / "Project_Planner_Blank.xlsx"), os.devnull], capture_output=True, text=True)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("Project start date is empty", result.stderr)


@unittest.skipUnless(SOFFICE, "LibreOffice not installed")
class ExportEditsTest(unittest.TestCase):
    """Edits made in the workbook, recalculated with LibreOffice, then exported."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.dir = Path(cls.tmp.name)
        cls.path = cls.dir / "Project_Planner.xlsx"
        build(cls.path)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def export(self, wb):
        out = self.dir / "edited.xlsx"
        wb.save(out)       # data_only workbook: values only, as the export reads them
        return tasks_by_uid(parse(mspdi.convert(out.read_bytes())))

    def test_baseline_and_ignored_links(self):
        first = recalc(self.path, self.dir)
        ws = first["Tasks"]
        col = columns(ws)
        snapshot = {(r, i): ws.cell(r, col[s]).value for r in range(FIRST, FIRST + 30) for i, s in enumerate(SOURCE)}

        def edit(wb):
            for (r, i), v in snapshot.items():
                wb["Tasks"].cell(r, col[BASELINE[i]]).value = v
            set_input(wb, 21, "Dur", 10)
            set_input(wb, 23, "P2", 30)            # a link to a task below: ignored, so not exported
        tasks = self.export(recalc(self.path, self.dir, edit))
        base = tasks[21].find(NS + "Baseline")
        self.assertEqual((text(base, "Number"), day(text(base, "Start")), day(text(base, "Finish"))),
                         ("0", dt.date(2026, 11, 10), dt.date(2026, 11, 19)))
        self.assertEqual((text(base, "Duration"), text(base, "Work"), text(base, "Cost")),
                         ("PT64H0M0S", "PT64H0M0S", "4160"))
        self.assertEqual(text(tasks[21], "Duration"), "PT80H0M0S")               # the current plan, not the baseline
        self.assertEqual([text(l, "PredecessorUID") for l in tasks[23].findall(NS + "PredecessorLink")], ["21"])


class ExportPageTest(unittest.TestCase):
    def test_page_is_up_to_date(self):
        page = (ROOT / "docs" / "export.html").read_text(encoding="utf-8")
        self.assertEqual(page, build_export_page.page(), "run: python scripts/build_export_page.py")
        self.assertNotIn("</script", page[page.index("const SOURCE"):page.index("const status")])

    @unittest.skipUnless(shutil.which("node") and os.environ.get("PYODIDE_DIR"),
                         "needs node with playwright, and PYODIDE_DIR (a local Pyodide 0.26.4)")
    def test_browser_writes_what_the_command_line_writes(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "page.xml"
            result = subprocess.run(["node", str(ROOT / "tests" / "browser" / "export_page.js"),
                                     str(ROOT / "docs" / "export.html"), str(SAMPLE), os.environ["PYODIDE_DIR"], str(out)],
                                    capture_output=True, text=True, timeout=600)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertIn("STATUS Done: 30 tasks, 7 resources.", result.stdout)
            self.assertEqual(out.read_text(encoding="utf-8"), mspdi.convert(SAMPLE.read_bytes()))
            blank = subprocess.run(["node", str(ROOT / "tests" / "browser" / "export_page.js"),
                                    str(ROOT / "docs" / "export.html"), str(ROOT / "Project_Planner_Blank.xlsx"),
                                    os.environ["PYODIDE_DIR"], str(out)], capture_output=True, text=True, timeout=600)
            self.assertIn("Project start date is empty", blank.stdout)


if __name__ == "__main__":
    unittest.main()
