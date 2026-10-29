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
from test_planner import CommittedWorkbooksTest, errors_in, named_cell, serial, table_rows  # noqa: E402
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
        self.assertTrue(rule.formula1.startswith("=INDEX(tsk_PickList,1):INDEX("), rule.formula1)
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

    def test_upgrade_keeps_engine_inputs(self):
        """B-110 / B-202 / B-201 / B-302 inputs survive an upgrade: Tasks columns, the setting, the Cost tables."""
        from planner import upgrade
        from test_release import put
        wb = make(baselined=True)
        put(wb, 28, {"Active": "No"})
        put(wb, 13, {"Remaining Duration": 7, "Lag1": "3ed"})
        named_cell(wb, "CompareBaseline").value = 2
        _put_in_table(wb["Cost"], "RateCard", 1, {"Per Use": 100})
        _put_in_table(wb["Cost"], "FixedCosts", 1, {"UID": 14, "Amount": 1000, "Description": "Licence"})
        buf = io.BytesIO()
        wb.save(buf)
        new = openpyxl.load_workbook(io.BytesIO(upgrade.convert(buf.getvalue())[0]))
        col = head(new["Tasks"])
        self.assertEqual(new["Tasks"].cell(task_row(new, 28), col["Active"]).value, "No")
        self.assertEqual(new["Tasks"].cell(task_row(new, 13), col["Remaining Duration"]).value, 7)
        self.assertEqual(new["Tasks"].cell(task_row(new, 13), col["Lag1"]).value, "3ed")
        self.assertEqual(named_cell(new, "CompareBaseline").value, 2)
        fixed = [r for r in table_rows(new, "Cost", "FixedCosts") if r["UID"]]
        self.assertEqual([(r["UID"], r["Amount"]) for r in fixed], [(14, 1000)])
        self.assertEqual(next(r for r in table_rows(new, "Cost", "RateCard") if r["Resource"])["Per Use"], 100)


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


def _plan_inputs_12(wb):
    """test_release._plan_inputs plus template 1.2's Active and Remaining Duration."""
    from test_release import _plan_inputs
    rows, settings, stages = _plan_inputs(wb)
    typed = [t for t in table_rows(wb, "Tasks", "Tasks")]
    for r, t in zip(rows, typed):
        r["Active"], r["RemDur"] = t["Active"], t["Remaining Duration"]
    return rows, settings, stages


def _put_in_table(ws, table, row, values):
    """Type values (header -> value) into data row `row` (1 = first) of a table on ws."""
    from openpyxl.utils.cell import range_boundaries
    c1, r1, c2, _ = range_boundaries(ws.tables[table].ref)
    col = {ws.cell(r1, c).value: c for c in range(c1, c2 + 1)}
    for h, v in values.items():
        ws.cell(r1 + row, col[h]).value = v


ENGINE_EDITS = {
    14: {"Lag1": "50%"},                                   # half of UID 13's 10 days
    21: {"Lag1": "3ed"},                                   # FS, elapsed
    23: {"Type1": "SS", "Lag1": "4ed"},
    18: {"Lag1": "2ed"},                                   # FF, elapsed
    25: {"Lag1": "2ed"},                                   # a milestone: on or after
    26: {"Type1": "SF", "Lag1": "12ed"},
    28: {"Active": "No"},                                  # inactive
    29: {"Pred2": 28},                                     # a link to it is ignored
    13: {"Remaining Duration": 7},                         # started: finishes 7 working days from the status date
    17: {"Remaining Duration": 3, "Actual Work": 40},
    16: {"Baseline 2 Start": None},
}


@unittest.skipUnless(__import__("test_planner").SOFFICE, "LibreOffice not installed")
class EngineTest(unittest.TestCase):
    """The v1.2 engine items, recalculated: % and elapsed lags (B-102), inactive tasks (B-110), Remaining Duration
    and Actual Work (B-202), baselines 2 and 3 (B-201), cost per use and fixed costs (B-302). The schedule is
    checked task by task against tests/reference_schedule.py (v12=True)."""

    @classmethod
    def setUpClass(cls):
        from test_planner import recalc
        from test_release import put
        cls.tmp = tempfile.TemporaryDirectory()
        path = Path(cls.tmp.name) / "engine.xlsx"
        make(baselined=True).save(path)
        cls.source = path

        def edit(wb):
            for uid, values in ENGINE_EDITS.items():
                put(wb, uid, {h: v for h, v in values.items() if v is not None})
            put(wb, 12, {"Baseline 2 Start": d(2026, 10, 1), "Baseline 2 Finish": d(2026, 10, 9)})
            cost = wb["Cost"]
            _put_in_table(cost, "RateCard", 1, {"Per Use": 100})
            _put_in_table(cost, "FixedCosts", 1, {"UID": 14, "Amount": 1000, "Description": "Gateway licence"})
            _put_in_table(cost, "FixedCosts", 2, {"UID": 14, "Amount": 500, "Description": "Sandbox account"})
            _put_in_table(cost, "FixedCosts", 3, {"UID": 13, "Amount": 300, "Description": "Test devices"})
        cls.wb = recalc(path, cls.tmp.name, edit)
        cls.t = tasks_by_uid(cls.wb)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()
        classic_again()

    def test_no_errors(self):
        self.assertEqual(errors_in(self.wb), [])

    def test_reference_schedule(self):
        import datetime as dt
        import reference_schedule
        from test_planner import serial, serial_or_none
        got, finish = reference_schedule.schedule(*_plan_inputs_12(self.wb), v12=True)
        bad = []
        for v in got:
            t = self.t[v["UID"]]
            for k, want in v.items():
                if k in ("UID", "CS", "CF"):
                    continue
                have = (t[k] == "Yes" if k == "Critical" else serial_or_none(t[k]) if isinstance(want, dt.date)
                        else t[k])
                if have != want:
                    bad.append((v["UID"], k, have, want))
        self.assertEqual(bad, [])
        self.assertEqual(serial(named_cell(self.wb, "ProjectFinish").value), finish)
        self.assertNotIn(28, [v["UID"] for v in got])

    def test_lags_as_typed(self):                                 # B-102
        self.assertEqual(self.t[14]["_Pred1Lag"], 5)
        self.assertEqual(self.t[21]["_Pred1Elapsed"], 3)
        self.assertIn("+3ed", self.t[21]["Predecessors by ID"] or "")
        self.assertEqual(self.t[21]["Issues"] or "", "")

    def test_inactive(self):                                      # B-110
        self.assertEqual(self.t[28]["Status"], "Inactive")
        self.assertIn(self.t[28]["Start"], ("", None))
        self.assertEqual(self.t[28]["_IsLeaf"], 0)
        self.assertIn("inactive", self.t[29]["Link Check"])

    def test_remaining_duration_and_actual_work(self):          # B-202
        t13 = self.t[13]
        self.assertEqual(serial(t13["Finish"]), serial(d(2026, 11, 3)))      # 7 working days from Mon 26-Oct
        self.assertAlmostEqual(t13["Progress"], 7 / 14, places=6)          # 15-23 Oct done (7 days), 7 to go
        t17 = self.t[17]
        self.assertAlmostEqual(t17["% Work Done"], 40 / t17["_LeafTotalWork"], places=6)
        self.assertAlmostEqual(self.t[17]["_LeafWorkDone"], 40 * (t17["_LeafWork"] / t17["_LeafTotalWork"]), places=6)

    def test_baseline_2(self):                                    # B-201
        self.assertEqual(serial(self.t[12]["Baseline 2 Start"]), serial(d(2026, 10, 1)))
        self.assertEqual(serial(self.t[12]["Planned Start"]), serial(self.t[12]["Baseline Start"]))   # compares 1

    def test_costs(self):                                         # B-302
        card = [r for r in table_rows(self.wb, "Cost", "RateCard") if r["Resource"]]
        first = card[0]["Resource"]
        cost = {r["ID"]: r for r in table_rows(self.wb, "Cost", "TaskCost") if r["ID"] not in (None, "")}
        for t in self.t.values():
            if t["_IsLeaf"] != 1:
                continue
            row = cost[t["ID"]]
            fixed = {14: 1500, 13: 300}.get(t["UID"], 0)
            uses = 100 if first in (t["Resource Names"] or "").split(", ") or (t["Resource Names"] or "").startswith(
                first + "[") or f", {first}[" in (t["Resource Names"] or "") else 0
            self.assertAlmostEqual(row["_LeafExtra"], uses + fixed, places=6, msg=t["Task Name"])
            started = t["_LeafProgress"] > 0
            self.assertAlmostEqual(row["_ExtraDone"], (uses if started else 0) + fixed * t["_LeafProgress"],
                                   places=6, msg=t["Task Name"])
        fixed_tasks = {r["UID"]: r["Task"] for r in table_rows(self.wb, "Cost", "FixedCosts") if r["UID"]}
        self.assertEqual(fixed_tasks[14], "Payment gateway integration")

    def test_ms_project_export(self):                             # B-102, B-110 in MSPDI
        import re
        from planner import mspdi
        path = Path(self.tmp.name) / "values.xlsx"
        self.wb.save(path)
        xml = mspdi.convert(path.read_bytes())
        xml = xml.decode() if isinstance(xml, bytes) else xml
        task = lambda uid: re.search(rf"<Task>\s*<UID>{uid}</UID>.*?</Task>", xml, re.S).group(0)
        self.assertIn("<Active>0</Active>", task(28))
        self.assertRegex(task(21), r"<LinkLag>43200</LinkLag>\s*<LagFormat>8</LagFormat>")      # 3 elapsed days
        self.assertIn("<Active>1</Active>", task(21))


@unittest.skipUnless(__import__("test_planner").SOFFICE, "LibreOffice not installed")
class CompareBaselineTest(unittest.TestCase):
    """B-201: Settings > CompareBaseline = 2 plans against Baseline 2 (and BaselineSet counts that baseline)."""

    @classmethod
    def tearDownClass(cls):
        classic_again()

    def test_compare_2(self):
        from test_planner import recalc, serial
        from test_release import put
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "b2.xlsx"
            make(baselined=True).save(path)

            def edit(wb):
                named_cell(wb, "CompareBaseline").value = 2
                put(wb, 12, {"Baseline 2 Start": d(2026, 10, 1), "Baseline 2 Finish": d(2026, 10, 9)})
            wb = recalc(path, tmp, edit)
        t = tasks_by_uid(wb)
        self.assertEqual(serial(t[12]["Planned Start"]), serial(d(2026, 10, 1)))
        self.assertEqual(serial(t[12]["Planned Finish"]), serial(d(2026, 10, 9)))
        self.assertEqual(serial(t[13]["Planned Start"]), serial(t[13]["_CalcStart"]))   # no baseline 2: current
        self.assertTrue(named_cell(wb, "BaselineSet").value)


class EngineStructureTest(unittest.TestCase):
    """Built, not recalculated: the MS Project export of the engine items and the TakeSnapshot links."""

    @classmethod
    def tearDownClass(cls):
        classic_again()

    def test_names_and_steps(self):
        wb = make(baselined=True)
        for n in ("tsk_BaseDates", "tsk_Baseline2", "tsk_Baseline3", "CompareBaseline", "cost_PerUse",
                  "cost_FixedUID", "cost_FixedAmount", "tsk_Active", "tsk_RemainingDuration"):
            self.assertIn(n, wb.defined_names, n)
        ws = wb["TakeSnapshot"]
        links = {c.hyperlink.location for row in ws.iter_rows() for c in row if c.hyperlink}
        self.assertTrue({"tsk_BaseDates", "tsk_Baseline2", "CompareBaseline"} <= links, links)



def d(y, m, day):
    import datetime as dt
    return dt.date(y, m, day)
