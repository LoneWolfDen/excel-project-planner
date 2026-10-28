"""Template 1.2 (the --release 1.2 build, release-1.2/): template 1.1 plus the Fluent design and the v1.2 backlog
items (docs/BACKLOG.md v1.2). The other editions don't change with any of this.

Run:  python -m unittest tests.test_release_12 -v      (the recalculation tests need LibreOffice)
"""
import io
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(ROOT))
import perf_check  # noqa: E402
from test_planner import CommittedWorkbooksTest, errors_in, named_cell  # noqa: E402
from test_release import head, small, task_row, tasks_by_uid  # noqa: E402

RELEASE = ROOT / "release-1.2"
SAMPLE, BLANK, SIMPLE = (RELEASE / n for n in ("Project_Planner.xlsx", "Project_Planner_Blank.xlsx",
                                               "Project_Planner_Simple.xlsx"))
V11 = ROOT / "release-1.1" / "Project_Planner.xlsx"


def make(**k):
    from planner.build import make_workbook
    return make_workbook(release_edition="1.2", **k)


def classic_again():
    from planner.build import make_workbook
    make_workbook(sizes=small())


def rgb(colour):
    return (colour.rgb or "")[-6:].upper() if colour is not None and isinstance(colour.rgb, str) else None


class DesignTest(unittest.TestCase):
    """The Fluent look: fonts, colours, tabs and pages."""

    @classmethod
    def setUpClass(cls):
        cls.wb = make(baselined=True)

    @classmethod
    def tearDownClass(cls):
        classic_again()

    def test_versions(self):
        self.assertEqual(self.wb.defined_names["TemplateVersion"].attr_text, '"1.2"')
        self.assertEqual(self.wb.defined_names["PlannerVersion"].attr_text, '"2.2"')

    def test_no_classic_colours_left(self):
        from planner.theme import COLOURS
        left = set()
        for ws in self.wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if not c.has_style:
                        continue
                    if c.font.name == "Arial":
                        left.add(f"{ws.title}!{c.coordinate} font Arial")
                    for colour in (c.font.color, c.fill.start_color if c.fill.fill_type == "solid" else None):
                        if rgb(colour) in COLOURS:
                            left.add(f"{ws.title}: {rgb(colour)}")
            for cf in ws.conditional_formatting:
                for rule in cf.rules:
                    d = rule.dxf
                    for colour in ((d.font.color if d and d.font else None),
                                   (d.fill.fgColor if d and d.fill else None), (d.fill.bgColor if d and d.fill else None)):
                        if rgb(colour) in COLOURS:
                            left.add(f"{ws.title} format: {rgb(colour)}")
        self.assertEqual(sorted(left)[:10], [])

    def test_meaning_kept(self):
        ws = self.wb["Tasks"]
        col = head(ws)
        self.assertEqual(rgb(ws.cell(6, col["Duration"]).font.color), "0F6CBD")         # input: brand blue
        self.assertEqual(rgb(ws.cell(6, col["Finish"]).fill.start_color), "F5F5F5")      # formula: grey
        self.assertEqual(rgb(ws.cell(5, col["Duration"]).fill.start_color), "0F548C")    # input header
        self.assertEqual(rgb(ws.cell(5, col["Finish"]).fill.start_color), "424242")      # formula header
        self.assertFalse(ws.cell(6, col["Duration"]).protection.locked)                 # still unlocked

    def test_pages_and_tabs(self):
        from planner.theme import PAGES, TABS
        for ws in self.wb.worksheets:
            self.assertEqual(rgb(ws.sheet_properties.tabColor), TABS[ws.title], ws.title)
            self.assertEqual(ws.sheet_view.showGridLines is False, ws.title in PAGES, ws.title)
        self.assertEqual(self.wb["Start"]["B1"].font.size, 24)

    def test_pick_list_shows_only_tasks_above(self):
        ws = self.wb["Tasks"]
        col = head(ws)
        rule = next(dv for dv in ws.data_validations.dataValidation
                    if openpyxl.utils.cell.range_boundaries(str(dv.sqref).split()[0])[0] == col["Pred1"])
        self.assertTrue(rule.formula1.startswith("=OFFSET("), rule.formula1)
        self.assertIn("COUNTIF(", rule.formula1)


class EditionsTest(unittest.TestCase):

    @classmethod
    def tearDownClass(cls):
        classic_again()

    def test_formula_rules(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "v12.xlsx"
            make(baselined=True).save(path)
            r = perf_check.static_report(path)
        self.assertEqual((r["a1_references"], r["missing_references"], r["excel_limit_problems"]), ([], [], []))
        self.assertEqual(r["volatile"], {"Tasks": {"OFFSET": 2200}})

    def test_other_editions_unchanged(self):
        from planner.build import make_workbook
        make(sizes=small())
        for edition in ("1.1", "m365"):
            wb = make_workbook(sizes=small(), release_edition=edition)
            self.assertEqual(wb["Tasks"]["A6"].font.name, "Arial", edition)
            self.assertNotIn("OFFSET(", next(dv.formula1 for dv in wb["Tasks"].data_validations.dataValidation
                                             if dv.formula1 and "PickList" in dv.formula1), edition)

    def test_modules(self):
        for combo in ((), ("check",), ("log",), ("check", "log", "history", "snapshot")):
            self.assertEqual(perf_check.missing_references(make(sizes=small(), modules=combo)), [], combo)


class CommittedTest(unittest.TestCase):
    compare = CommittedWorkbooksTest.compare

    def test_sample(self):
        self.compare(SAMPLE, "--release", "1.2", "--baselined")

    def test_blank(self):
        self.compare(BLANK, "--release", "1.2", "--blank")

    def test_simple(self):
        self.compare(SIMPLE, "--release", "1.2", "--baselined", "--simple")


class CommittedSampleTest(unittest.TestCase):

    def test_no_errors(self):
        for path in (SAMPLE, BLANK, SIMPLE):
            self.assertEqual(errors_in(openpyxl.load_workbook(path, data_only=True)), [], path.name)

    def test_same_plan_as_11(self):
        new = tasks_by_uid(openpyxl.load_workbook(SAMPLE, data_only=True))
        old = tasks_by_uid(openpyxl.load_workbook(V11, data_only=True))
        keys = ("Start", "Finish", "Total Slack", "Critical", "Work", "Issues")
        self.assertEqual([(u, k) for u in old for k in keys if old[u][k] != new[u][k]], [])


class ToolsTest(unittest.TestCase):

    @classmethod
    def tearDownClass(cls):
        classic_again()

    def test_upgrade_keeps_typed_inputs(self):
        """The upgrade finds input columns by their header colour: template 1.2's must count too."""
        from planner import upgrade
        wb = openpyxl.load_workbook(SAMPLE)
        ws = wb["Log"]
        col = head(ws, 4)
        free = next(r for r in range(5, 160) if not ws.cell(r, col["Title"]).value)
        ws.cell(free, col["Title"]).value, ws.cell(free, col["Type"]).value = "Vendor contract signed", "Decision"
        buf = io.BytesIO()
        wb.save(buf)
        new = openpyxl.load_workbook(io.BytesIO(upgrade.convert(buf.getvalue())[0]))
        self.assertEqual(new.defined_names["PlannerVersion"].attr_text, '"2.2"')
        titles = [new["Log"].cell(r, col["Title"]).value for r in range(5, 160)]
        self.assertIn("Vendor contract signed", titles)
        self.assertEqual(new["Tasks"].cell(task_row(new, 17), head(new["Tasks"])["Pred2"]).value, 13)


SAMPLES = RELEASE / "samples"


class UXTest(unittest.TestCase):
    """The v1.2 UX items, on the committed sample (recalculated by the build)."""

    @classmethod
    def setUpClass(cls):
        cls.wb = openpyxl.load_workbook(SAMPLE, data_only=True)
        cls.t = tasks_by_uid(cls.wb)

    def test_msp_status(self):                                   # B-206
        allowed = {"Complete", "On Schedule", "Late", "Future Task"}
        self.assertTrue({r["MSP Status"] for r in self.t.values()} <= allowed)
        for r in self.t.values():
            if r["Progress"] == 1:
                self.assertEqual(r["MSP Status"], "Complete", r["Task Name"])
            elif r["Status"] == "Behind":
                self.assertEqual(r["MSP Status"], "Late", r["Task Name"])
        self.assertEqual(self.t[30]["MSP Status"], "Future Task")

    def test_board(self):                                         # B-404
        from planner.tasks import BOARD_STATUSES
        ws = self.wb["Board"]
        leaves = [r for r in self.t.values() if r["_IsLeaf"] == 1]
        for s, status in enumerate(BOARD_STATUSES, start=1):
            col = 2 + 2 * s
            head = ws.cell(7, col).value
            want = sorted(r["Task Name"] for r in leaves if r["Status"] == status)
            self.assertTrue(head.startswith(f"{status}  ({len(want)})"), head)
            cards = sorted(str(ws.cell(r, col).value).split("\n")[0] for r in range(8, 33) if ws.cell(r, col).value)
            self.assertEqual(cards, want, status)

    def test_backlog(self):                                       # B-405
        rows = {r["Item"]: r for r in table_rows(self.wb, "Backlog", "ProductBacklog") if r["Item"]}
        self.assertEqual(rows["Accessibility review"]["On Tasks"], "Yes")
        self.assertIn(rows["Dark mode"]["On Tasks"], ("", None))
        totals = {self.wb["Backlog"].cell(r, 2).value: self.wb["Backlog"].cell(r, 3).value for r in range(5, 10)}
        self.assertEqual(totals["Must"], 11)

    def test_locale_safe_dates(self):                             # B-010
        import re
        wb = openpyxl.load_workbook(SAMPLE)
        dated = [f"{ws.title}!{c.coordinate}" for ws in wb.worksheets for row in ws.iter_rows() for c in row
                 if isinstance(c.value, str) and re.search(r'TEXT\([^"]*"[dmy]', c.value)]
        self.assertEqual(dated[:5], [])
        self.assertIn("2027", named_cell(self.wb, "sts_Narrative").value)        # still a readable date
        self.assertIn("-Jan-2027", named_cell(self.wb, "sts_Narrative").value)

    def test_summary_line_and_printing(self):                     # B-111, B-508
        ws = self.wb["Tasks"]
        self.assertTrue(str(ws["A3"].value).startswith("Project 07-Sep-26 to "), ws["A3"].value)
        ws = openpyxl.load_workbook(SAMPLE)["Tasks"]
        self.assertEqual((ws.page_setup.paperSize, ws.page_setup.orientation), (8, "landscape"))
        self.assertTrue(ws.print_title_rows)

    def test_custom_fields_go_to_ms_project(self):               # B-509
        from planner import mspdi
        xml = mspdi.convert(SAMPLE.read_bytes())
        xml = xml.decode() if isinstance(xml, bytes) else xml
        self.assertIn("<FieldName>Text2</FieldName>", xml)
        self.assertIn("<Alias>Text 1</Alias>", xml)
        self.assertIn("<FieldName>Flag1</FieldName>", xml)
        self.assertIn("<Value>Payments</Value>", xml)

    def test_exports(self):                                       # B-506, B-706
        import csv
        from planner.exports import network, to_csv
        mmd = network(SAMPLE.read_bytes())
        self.assertTrue(mmd.startswith("flowchart LR"))
        self.assertIn("t16 --> t17", mmd)
        self.assertIn("t13 -->|SS+2d| t17", mmd)
        self.assertIn("class ", mmd)                             # critical tasks highlighted
        self.assertIn("digraph plan", network(SAMPLE.read_bytes(), "dot"))
        for fmt, column in (("jira", "Summary"), ("azure", "Title"), ("planner", "Task Name")):
            rows = list(csv.DictReader(io.StringIO(to_csv(SAMPLE.read_bytes(), fmt))))
            self.assertIn("Core API endpoints", [r[column] for r in rows], fmt)
        jira = list(csv.DictReader(io.StringIO(to_csv(SAMPLE.read_bytes(), "jira"))))
        self.assertEqual(next(r for r in jira if r["Summary"] == "Core API endpoints")["Parent Id"], "11")

    def test_more_samples(self):                                  # B-804
        for name, finish in (("Construction_Planner.xlsx", None), ("Event_Planner.xlsx", None)):
            wb = openpyxl.load_workbook(SAMPLES / name, data_only=True)
            self.assertEqual(errors_in(wb), [], name)
            self.assertGreater(len(tasks_by_uid(wb)), 15, name)
        event = tasks_by_uid(openpyxl.load_workbook(SAMPLES / "Event_Planner.xlsx", data_only=True))
        self.assertEqual(str(event[19]["Start"])[:10], "2027-01-27")
        self.assertEqual(named_cell(openpyxl.load_workbook(SAMPLES / "Construction_Planner.xlsx", data_only=True),
                                    "ProjectType").value, "Construction")


@unittest.skipUnless(__import__("test_planner").SOFFICE, "LibreOffice not installed")
class BoardFilterTest(unittest.TestCase):

    @classmethod
    def tearDownClass(cls):
        classic_again()

    def test_person_filter(self):
        from test_planner import recalc
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "v12.xlsx"
            make(baselined=True).save(path)
            wb = recalc(path, tmp, lambda wb: setattr(named_cell(wb, "BoardResource"), "value", "Jordan Lee"))
        ws = wb["Board"]
        cards = [str(ws.cell(r, c).value) for c in range(4, 13, 2) for r in range(8, 33) if ws.cell(r, c).value]
        self.assertTrue(cards)
        self.assertTrue(all("Jordan Lee" in c for c in cards), cards)


class SampleBuildTest(unittest.TestCase):
    compare = CommittedWorkbooksTest.compare

    def test_construction(self):
        self.compare(SAMPLES / "Construction_Planner.xlsx", "--release", "1.2", "--sample", "construction")

    def test_event(self):
        self.compare(SAMPLES / "Event_Planner.xlsx", "--release", "1.2", "--sample", "event")


if __name__ == "__main__":
    unittest.main()
