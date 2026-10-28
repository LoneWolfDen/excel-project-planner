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


if __name__ == "__main__":
    unittest.main()
