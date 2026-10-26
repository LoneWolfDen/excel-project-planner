"""Tests for the Excel project planner.

Builds the workbook with scripts/build_planner.py, recalculates it with LibreOffice
(headless) and checks the scheduling results against hand-worked dates.

Run:  python -m unittest discover -s tests -v
Needs openpyxl and LibreOffice (`soffice` on PATH); recalculation tests are skipped without it.
"""
import datetime as dt
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
BUILD = ROOT / "scripts" / "build_planner.py"
SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")
STATUS_DATE = dt.date(2026, 10, 25)
ERRORS = ("#NAME?", "#VALUE!", "#REF!", "#DIV/0!", "#N/A", "#NUM!", "#NULL!", "Err:")
FIRST = 6  # first task row on Tasks

# Tasks headers used by the tests (columns are found by header, not letter)
H = {
    "ID": "ID", "UID": "UID", "WBS": "WBS", "Level": "Level", "Name": "Task Name", "Dur": "Duration",
    "P1": "Pred1 UID", "P2": "Pred2 UID", "AS": "Actual Start", "AF": "Actual Finish", "Pct": "% Complete",
    "PS": "Planned Start", "PF": "Planned Finish", "FS": "Forecast Start", "FF": "Forecast Finish",
    "SVar": "Start Var", "Var": "Finish Var", "DLChk": "Deadline Check", "Link": "Link Check", "Prog": "Progress",
    "Status": "Status", "Work": "Work", "Cost": "Cost", "Sprint": "Sprint", "Pred": "Predecessors",
    "DVar": "Duration Var", "WVar": "Work Var", "CVar": "Cost Var",
}
SOURCE = ["BL Source Start", "BL Source Finish", "BL Source Duration", "BL Source Work", "BL Source Cost"]
BASELINE = ["Baseline Start", "Baseline Finish", "Baseline Duration", "Baseline Work", "Baseline Cost"]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import perf_check  # noqa: E402


def build(path, *flags):
    subprocess.run([sys.executable, str(BUILD), str(path), *flags], check=True, capture_output=True)


def columns(ws):
    return {ws.cell(5, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(5, c).value}


def named_cell(wb, name):
    sheet, ref = wb.defined_names[name].attr_text.split("!")
    return wb[sheet.strip("'")][ref.replace("$", "")]


def recalc(src, outdir, edit=None, status_date=True):
    """Optionally edit inputs (edit(wb) with formulas loaded), then let LibreOffice recalculate a copy."""
    wb = openpyxl.load_workbook(src)
    if status_date:
        named_cell(wb, "StatusDateInput").value = STATUS_DATE   # deterministic results
    if edit:
        edit(wb)
    stage = Path(tempfile.mkdtemp(dir=outdir))
    staged = stage / src.name
    wb.save(staged)
    subprocess.run([SOFFICE, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(stage / "out"), str(staged)],
                   check=True, capture_output=True, timeout=900)
    return openpyxl.load_workbook(stage / "out" / src.name, data_only=True)


def d(value):
    return value.date() if isinstance(value, dt.datetime) else value


class Tasks:
    """Read a recalculated Tasks sheet by task ID and header key."""

    def __init__(self, wb):
        self.ws = wb["Tasks"]
        self.col = columns(self.ws)

    def row_of(self, task_id):
        c = self.col[H["ID"]]
        for r in range(FIRST, self.ws.max_row + 1):
            if self.ws.cell(r, c).value == task_id:
                return r
        raise KeyError(task_id)

    def get(self, task_id, key):
        return d(self.ws.cell(self.row_of(task_id), self.col[H.get(key, key)]).value)


def set_input(wb, task_id, key, value):
    """Set an input cell on the unrecalculated workbook (sample task n is on row n + 5)."""
    ws = wb["Tasks"]
    ws.cell(task_id + FIRST - 1, columns(ws)[H.get(key, key)]).value = value


def errors_in(wb):
    return [f"{ws.title}!{c.coordinate}={c.value}" for ws in wb.worksheets for r in ws.iter_rows() for c in r
            if isinstance(c.value, str) and c.value.startswith(ERRORS)]


class Tmp(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.dir = Path(cls.tmp.name)
        cls.path = cls.dir / "Project_Planner.xlsx"
        build(cls.path)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()


class StructureTest(Tmp):
    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.wb = openpyxl.load_workbook(cls.path)

    def test_sheets(self):
        self.assertEqual(self.wb.sheetnames, ["Instructions", "Dashboard", "Settings", "Tasks", "Gantt", "Resources",
                                              "Sprints", "Holidays", "MSP_Tasks", "MSP_Resources", "Lists"])

    def test_named_ranges(self):
        names = set(self.wb.defined_names)
        for n in ("ProjectStart", "ProjectStartDate", "StatusDate", "WorkWeek", "Holidays", "HoursPerDay", "GanttStart",
                  "GanttPeriod", "Sprint1Start", "SprintLength", "NextFreeUID", "tsk_ID", "tsk_UID", "tsk_PlanStart",
                  "tsk_CalcStart", "tsk_FcFinish", "tsk_Resource", "tsk_Deadline", "tsk_DeadlineCheck", "tsk_LinkCheck",
                  "tsk_BaseStart", "tsk_BaseCost", "tsk_CostVariance", "res_Name", "res_Rate", "spr_Name", "spr_Start",
                  "gantt_Origin", "load_Grid"):
            self.assertIn(n, names)
        # column names are structured references, so they grow with the table
        self.assertEqual(self.wb.defined_names["tsk_Duration"].attr_text, "Tasks[Duration]")
        self.assertEqual(self.wb.defined_names["Holidays"].attr_text, "HolidayTable[Date]")

    def test_tables_have_calculated_columns(self):
        tasks = self.wb["Tasks"].tables["Tasks"]
        calculated = {c.name for c in tasks.tableColumns if c.calculatedColumnFormula is not None}
        self.assertIn("Forecast Finish", calculated)
        self.assertNotIn("Duration", calculated)                 # inputs are not calculated columns
        for sheet, table in (("Resources", "ResourceTable"), ("Sprints", "SprintTable"), ("Gantt", "GanttView"),
                             ("MSP_Tasks", "MSPTasks"), ("Holidays", "HolidayTable")):
            self.assertIn(table, self.wb[sheet].tables)

    def test_no_a1_references_in_formulas(self):
        report = perf_check.static_report(self.path)
        self.assertEqual(report["a1_references"], [])

    def test_formulas_within_excel_limits(self):
        report = perf_check.static_report(self.path)
        self.assertEqual(report["excel_limit_problems"], [])
        self.assertEqual(perf_check.excel_limit_problems('=IF(A,"(",B'), ["unbalanced parentheses, quotes or brackets"])
        self.assertEqual(perf_check.excel_limit_problems('=IF(Tasks[[#This Row],[x)]]=1,"a)",2)'), [])

    def test_a1_lint_catches_cell_references(self):
        names = {"ProjectStart"}
        self.assertEqual(perf_check.a1_references("=SUM(A1:B2)+Settings!$C$5+ProjectStart", names), ["A1:B2", "Settings!$C$5"])
        self.assertEqual(perf_check.a1_references("=Tasks[[#This Row],[Duration]]*2", names), [])

    def test_volatile_functions_only_where_allowed(self):
        # OFFSET only in Tasks, for the few columns that read another row's calculated value; no TODAY, INDIRECT ...
        volatile = perf_check.static_report(self.path)["volatile"]
        self.assertEqual(volatile, {"Tasks": {"OFFSET": 1000}})
        ws = self.wb["Tasks"]
        offset_cols = {ws.cell(5, c).value for c in range(1, ws.max_column + 1)
                       if "OFFSET(" in str(ws.cell(FIRST, c).value)}
        self.assertEqual(offset_cols, {"WBS", "_Pred1PlanDriver", "_Pred2PlanDriver", "_Pred1FcstDriver", "_Pred2FcstDriver"})

    def test_gantt_timeline_has_no_formulas(self):
        ws = self.wb["Gantt"]
        origin = named_cell(self.wb, "gantt_Origin")
        for row in ws.iter_rows(min_row=origin.row + 1, max_row=origin.row + 30, min_col=origin.column + 1,
                                max_col=origin.column + 150):
            for c in row:
                self.assertIsNone(c.value, c.coordinate)

    def test_scheduling_formulas_use_names(self):
        ws = self.wb["Tasks"]
        drivers = [ws.cell(7, c).value for c in range(1, ws.max_column + 1)
                   if isinstance(ws.cell(7, c).value, str) and "CHOOSE(" in ws.cell(7, c).value]
        self.assertTrue(drivers)
        for token in ("Tasks[[#Headers],[_LeafCalcFinish]]", "Tasks[[#This Row],[_Pred1Pos]]", "WorkWeek", "Holidays"):
            self.assertIn(token, drivers[0])
        # every row of a calculated column carries the same formula text: nothing points at a particular row
        for name in ("Forecast Start", "_Pred1PlanDriver", "WBS"):
            col = columns(ws)[name]
            self.assertEqual(len({ws.cell(r, col).value for r in range(FIRST, FIRST + 200)}), 1, name)

    def test_uids_prefilled(self):
        ws = self.wb["Tasks"]
        c = columns(ws)["UID"]
        self.assertEqual([ws.cell(r, c).value for r in range(FIRST, FIRST + 200)], list(range(1, 201)))

    def test_views_read_tasks_by_name_not_row(self):
        # Gantt and MSP_Tasks use INDEX(tsk_..., n), so rows inserted on Tasks still show up
        for sheet in ("Gantt", "MSP_Tasks"):
            for row in self.wb[sheet].iter_rows(max_row=30, max_col=25):
                for c in row:
                    if isinstance(c.value, str) and c.value.startswith("="):
                        self.assertNotIn("Tasks!", c.value, f"{sheet}!{c.coordinate}")

    def test_msp_headers_match_project_fields(self):
        headers = [c.value for c in self.wb["MSP_Tasks"][1]]
        for h in ("ID", "Unique ID", "Name", "Outline Level", "Duration", "Predecessors", "Unique ID Predecessors",
                  "Resource Names", "% Complete", "Actual Start", "Actual Finish", "Baseline Start", "Baseline Finish",
                  "Baseline Duration", "Baseline Work", "Baseline Cost", "Deadline"):
            self.assertIn(h, headers)

    def test_no_post_2010_functions(self):
        banned = ("XLOOKUP", "FILTER(", "SEQUENCE(", "UNIQUE(", "LET(", "_xlfn.")
        for ws in self.wb.worksheets:
            for r in ws.iter_rows(max_row=12):
                for c in r:
                    if isinstance(c.value, str) and c.value.startswith("="):
                        for b in banned:
                            self.assertNotIn(b, c.value, f"{ws.title}!{c.coordinate}")


class CommittedWorkbooksTest(unittest.TestCase):
    """The .xlsx files in the repo must match what the build script produces (run the build commands in README)."""

    def compare(self, committed, *flags):
        with tempfile.TemporaryDirectory() as tmp:
            fresh = Path(tmp) / committed.name
            build(fresh, *flags)
            a, b = openpyxl.load_workbook(fresh), openpyxl.load_workbook(committed)
            self.assertEqual(a.sheetnames, b.sheetnames)
            diffs = []
            for ws in a.worksheets:
                other = b[ws.title]
                for row in ws.iter_rows():
                    for c in row:
                        x, y = d(c.value), d(other[c.coordinate].value)
                        if x != y and not (x in (None, "") and y in (None, "")):
                            diffs.append(f"{ws.title}!{c.coordinate}: build={x!r} committed={y!r}")
            self.assertEqual(diffs[:10], [], f"{committed.name} is out of date ({len(diffs)} cells differ)")
            # LibreOffice's save drops calculated-column formulas; the build restores them so Excel auto-fills rows
            with zipfile.ZipFile(committed) as z:
                tables = [z.read(n).decode() for n in z.namelist() if n.startswith("xl/tables/")]
            tasks = next(t for t in tables if 'displayName="Tasks"' in t)
            self.assertIn("<calculatedColumnFormula>", tasks)

    def test_sample_workbook_up_to_date(self):
        self.compare(ROOT / "Project_Planner.xlsx")

    def test_blank_workbook_up_to_date(self):
        self.compare(ROOT / "Project_Planner_Blank.xlsx", "--blank")


@unittest.skipUnless(SOFFICE, "LibreOffice not installed")
class ScheduleTest(Tmp):
    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.wb = recalc(cls.path, cls.dir)
        cls.t = Tasks(cls.wb)

    def val(self, task_id, key):
        return self.t.get(task_id, key)

    def test_no_formula_errors(self):
        self.assertEqual(errors_in(self.wb), [])

    def test_wbs_and_outline(self):
        self.assertEqual(self.val(1, "WBS"), "1")
        self.assertEqual(self.val(13, "WBS"), "3.1.2")
        self.assertEqual(self.val(30, "WBS"), "5.3")

    def test_finish_to_start(self):
        self.assertEqual(self.val(2, "PS"), dt.date(2026, 9, 7))
        self.assertEqual(self.val(2, "PF"), dt.date(2026, 9, 9))
        self.assertEqual(self.val(3, "PS"), dt.date(2026, 9, 10))

    def test_start_to_start_with_lag(self):
        self.assertEqual(self.val(7, "PS"), dt.date(2026, 9, 17))   # task 6 starts Mon 14-Sep, SS+3

    def test_finish_to_finish_with_lag(self):
        self.assertEqual(self.val(18, "PF"), dt.date(2026, 10, 29))  # task 17 finishes 28-Oct, FF+1
        self.assertEqual(self.val(18, "PS"), dt.date(2026, 10, 27))

    def test_two_predecessors_take_latest(self):
        self.assertEqual(self.val(9, "PS"), dt.date(2026, 10, 1))   # max(task 7, task 8 finish)

    def test_weekend_skipped(self):
        self.assertEqual(self.val(6, "PF"), dt.date(2026, 9, 23))   # 8 days from Mon 14-Sep

    def test_holidays_skipped(self):
        # Hypercare: 10 days from Wed 16-Dec skipping 24, 25, 31 Dec and 1 Jan
        self.assertEqual(self.val(26, "PF"), dt.date(2027, 1, 4))

    def test_start_no_earlier_than(self):
        self.assertEqual(self.val(22, "PS"), dt.date(2026, 12, 7))

    def test_manual_task(self):
        self.assertEqual(self.val(28, "PS"), dt.date(2026, 11, 16))

    def test_summary_rollup(self):
        self.assertEqual(self.val(10, "PS"), dt.date(2026, 10, 2))
        self.assertEqual(self.val(10, "PF"), dt.date(2026, 10, 30))
        self.assertEqual(self.val(11, "Work"), 168)   # 40 + 80 + 48 hrs
        self.assertAlmostEqual(self.val(1, "Prog"), 1.0)

    def test_actuals_drive_forecast(self):
        self.assertEqual(self.val(2, "FF"), dt.date(2026, 9, 10))   # actual finish used
        self.assertEqual(self.val(13, "FF"), dt.date(2026, 10, 30))  # in progress, 50% of 10 days left after status
        self.assertEqual(self.val(14, "FS"), dt.date(2026, 11, 2))   # successor pushed
        self.assertEqual(self.val(13, "Status"), "Behind")
        self.assertEqual(self.val(14, "Status"), "Late Start")
        self.assertEqual(self.val(10, "Var"), 6)
        self.assertEqual(self.val(14, "SVar"), 6)

    def test_predecessor_text(self):
        self.assertEqual(self.val(7, "Pred"), "6SS+3d")
        self.assertEqual(self.val(17, "Pred"), "16,13SS+2d")
        self.assertEqual(self.val(18, "Pred"), "17FF+1d")

    def test_deadlines(self):
        self.assertEqual(self.val(19, "DLChk"), "Late by 1d")   # forecast Mon 09-Nov vs deadline Fri 06-Nov
        self.assertEqual(self.val(25, "DLChk"), "OK")
        self.assertIsNone(self.val(2, "DLChk"))
        self.assertEqual(self.dashboard("Missed deadlines"), 1)
        gantt = self.wb["Gantt"]
        col = {c.value: c.column for c in gantt[5] if c.value}
        self.assertEqual(d(gantt.cell(self.t.row_of(19), col["_Deadline"]).value), dt.date(2026, 11, 6))

    def test_no_link_problems_in_sample(self):
        self.assertEqual(self.dashboard("Link problems (see Tasks > Link Check)"), 0)

    def test_baseline_not_set_by_default(self):
        self.assertEqual(self.dashboard("Baseline"), "Not set (planned = calculated)")
        self.assertIsNone(self.val(2, "DVar"))

    def test_sprint_and_cost(self):
        self.assertEqual(self.val(2, "Sprint"), "Sprint 1")
        self.assertEqual(self.val(2, "Cost"), 24 * 95)
        self.assertEqual(self.dashboard("Current sprint"), "Sprint 4")

    def test_msp_export(self):
        mx = self.wb["MSP_Tasks"]
        col = {c.value: c.column for c in mx[1]}
        row = lambda task_id: task_id + 1
        self.assertEqual(mx.cell(row(6), col["Duration"]).value, "8d")
        self.assertEqual(mx.cell(row(18), col["Resource Names"]).value, "Casey Kim[50%]")
        self.assertEqual(mx.cell(row(22), col["Constraint Type"]).value, "Start No Earlier Than")
        self.assertEqual(mx.cell(row(19), col["Unique ID Predecessors"]).value, "14,18")
        self.assertEqual(d(mx.cell(row(19), col["Deadline"]).value), dt.date(2026, 11, 6))

    def dashboard(self, label):
        ws = self.wb["Dashboard"]
        for r in range(1, 40):
            if ws.cell(r, 2).value == label:
                return ws.cell(r, 3).value
        raise KeyError(label)


@unittest.skipUnless(SOFFICE, "LibreOffice not installed")
class BaselineTest(Tmp):
    """Frozen baseline: copy the source block to the baseline block (values), then change the plan."""

    def test_baseline_freezes_plan_and_measures_variance(self):
        first = recalc(self.path, self.dir)
        ws1 = first["Tasks"]
        col = columns(ws1)
        snapshot = {(r, i): ws1.cell(r, col[s]).value for r in range(FIRST, FIRST + 30) for i, s in enumerate(SOURCE)}

        def set_baseline_then_slip(wb):
            ws = wb["Tasks"]
            for (r, i), v in snapshot.items():
                ws.cell(r, col[BASELINE[i]]).value = v      # what Paste Special > Values does
            set_input(wb, 2, "Dur", 5)                      # charter now takes 5 days instead of 3

        wb = recalc(self.path, self.dir, set_baseline_then_slip)
        self.assertEqual(errors_in(wb), [])
        t = Tasks(wb)
        # planned dates stay on the baseline (baseline taken from the forecast, which uses actuals)
        self.assertEqual(t.get(3, "PS"), dt.date(2026, 9, 11))
        self.assertEqual(t.get(2, "PF"), dt.date(2026, 9, 10))
        # variances against the baseline
        self.assertEqual(t.get(2, "DVar"), 2)
        self.assertEqual(t.get(2, "WVar"), 16)
        self.assertEqual(t.get(2, "CVar"), 16 * 95)
        self.assertEqual(t.get(1, "DVar"), 0)       # summary duration from dates; actual finish unchanged
        self.assertEqual(t.get(3, "Var"), 0)        # completed tasks keep their actual dates
        dash = {wb["Dashboard"].cell(r, 2).value: wb["Dashboard"].cell(r, 3).value for r in range(5, 20)}
        self.assertEqual(dash["Baseline"], "Set on 30 tasks")

    def test_unstarted_task_slip_shows_against_baseline(self):
        first = recalc(self.path, self.dir)
        ws1 = first["Tasks"]
        col = columns(ws1)
        snapshot = {(r, i): ws1.cell(r, col[s]).value for r in range(FIRST, FIRST + 30) for i, s in enumerate(SOURCE)}

        def edit(wb):
            ws = wb["Tasks"]
            for (r, i), v in snapshot.items():
                ws.cell(r, col[BASELINE[i]]).value = v
            set_input(wb, 21, "Dur", 10)   # system testing grows from 8 to 10 days

        t = Tasks(recalc(self.path, self.dir, edit))
        self.assertEqual(t.get(21, "PF"), dt.date(2026, 11, 19))   # frozen
        self.assertEqual(t.get(21, "FF"), dt.date(2026, 11, 23))
        self.assertEqual(t.get(21, "Var"), 2)
        self.assertEqual(t.get(22, "Var"), 0)                       # UAT's Start No Earlier Than absorbs the slip


@unittest.skipUnless(SOFFICE, "LibreOffice not installed")
class LinkCheckTest(Tmp):
    def test_bad_links_are_explained_not_circular(self):
        def edit(wb):
            set_input(wb, 7, "P1", 7)      # itself
            set_input(wb, 13, "P2", 11)    # its own summary (Backend)
            set_input(wb, 14, "P1", 999)   # missing
            set_input(wb, 6, "P2", 150)    # empty row (below, too)
            set_input(wb, 8, "P2", 21)     # a task further down: predecessors must be above
            set_input(wb, 21, "P2", 15)    # another phase's summary (Frontend)

        wb = recalc(self.path, self.dir, edit)
        self.assertEqual(errors_in(wb), [])
        t = Tasks(wb)
        self.assertEqual(t.get(7, "Link"), "Pred 1: links to itself.")
        self.assertEqual(t.get(13, "Link"), "Pred 2: UID 11 is a summary: link to one of its tasks.")
        self.assertEqual(t.get(14, "Link"), "Pred 1: UID 999 not found.")
        self.assertEqual(t.get(6, "Link"), "Pred 2: UID 150 is an empty row.")
        self.assertEqual(t.get(8, "Link"), "Pred 2: UID 21 is below this task: move it above.")
        self.assertEqual(t.get(21, "Link"), "Pred 2: UID 15 is a summary: link to one of its tasks.")
        self.assertEqual(t.get(8, "PS"), dt.date(2026, 9, 24))     # the valid link still applies
        # task 14 lost its link (UID 999), so the build milestone moves to 29-Oct (task 18) and testing follows it
        self.assertEqual(t.get(21, "PS"), dt.date(2026, 10, 30))
        self.assertEqual(t.get(7, "PS"), dt.date(2026, 9, 7))      # bad link ignored -> project start
        self.assertEqual(t.get(13, "PS"), dt.date(2026, 10, 9))    # good link still applies
        self.assertIn(t.get(7, "Pred"), (None, ""))
        dash = {wb["Dashboard"].cell(r, 2).value: wb["Dashboard"].cell(r, 3).value for r in range(5, 20)}
        self.assertEqual(dash["Link problems (see Tasks > Link Check)"], 6)


@unittest.skipUnless(SOFFICE, "LibreOffice not installed")
class InsertRowTest(Tmp):
    """Inserting a task mid-list renumbers IDs but links follow UIDs.

    Every task row carries identical relative formulas, so shifting the input cells down one row is exactly what
    Excel's 'Insert Copied Cells' of a blank task row produces.
    """

    def test_insert_task_keeps_links(self):
        def insert(wb):
            ws = wb["Tasks"]
            inputs = [c for c in range(1, ws.max_column + 1) if (ws.cell(5, c).fill.start_color.rgb or "").endswith("1F4E78")]
            at = 14 + FIRST - 1                          # insert above "Payment gateway integration"
            last = FIRST + 199
            for r in range(last - 1, at - 1, -1):
                for c in inputs:
                    ws.cell(r + 1, c).value = ws.cell(r, c).value
            for c in inputs:
                ws.cell(at, c).value = None
            new = {"UID": 201, "Level": 3, "Name": "Security review", "Dur": 2, "P1": 13}
            col = columns(ws)
            for k, v in new.items():
                ws.cell(at, col[H[k]]).value = v

        wb = recalc(self.path, self.dir, insert)
        self.assertEqual(errors_in(wb), [])
        t = Tasks(wb)
        self.assertEqual(t.get(14, "Name"), "Security review")
        self.assertEqual(t.get(14, "WBS"), "3.1.3")
        self.assertEqual(t.get(14, "PS"), dt.date(2026, 10, 23))
        self.assertEqual(t.get(15, "Name"), "Payment gateway integration")
        self.assertEqual(t.get(15, "WBS"), "3.1.4")
        self.assertEqual(t.get(15, "PS"), dt.date(2026, 10, 23))     # still follows UID 13
        self.assertEqual(t.get(20, "Name"), "Build complete")
        self.assertEqual(t.get(20, "Pred"), "15,19")                  # by ID: shifted
        self.assertEqual(t.get(18, "Pred"), "17,13SS+2d")             # Feature screens now ID 18
        self.assertEqual(t.get(11, "Work"), 168 + 0)                  # no resource on new task
        gantt = wb["Gantt"]
        self.assertEqual(gantt.cell(14 + FIRST - 1, 3).value.strip(), "Security review")


import lo_session  # noqa: E402


@unittest.skipUnless(SOFFICE and lo_session.AVAILABLE, "LibreOffice / python3-uno not installed")
class RealRowEditTest(Tmp):
    """Insert and delete real rows, as a person would in Excel, and check every existing formula still works.

    This is what catches formulas that only look robust: a range that ends at the row above, for example, does
    not grow when a row is inserted directly above, so the row below silently loses sight of the new task.
    """

    def edit(self, fn):
        prepared = self.dir / f"prepared_{fn.__name__}.xlsx"
        wb = openpyxl.load_workbook(self.path)
        named_cell(wb, "StatusDateInput").value = STATUS_DATE
        wb.save(prepared)
        out = self.dir / f"edited_{fn.__name__}.xlsx"
        cols = columns(wb["Tasks"])
        with lo_session.Calc(prepared) as calc:
            fn(calc, cols)
            calc.save(out)
        return openpyxl.load_workbook(out, data_only=True)

    def test_insert_task_directly_above_and_link_to_it(self):
        def insert(calc, cols):
            row = 14 + FIRST - 1                                    # "Payment gateway integration"
            calc.insert_row_like_excel("Tasks", row, max(cols.values()))
            for c in cols.values():                                  # clear the copied inputs, keep the formulas
                if not calc.sheet("Tasks").getCellByPosition(c - 1, row - 1).getFormula().startswith("="):
                    calc.set("Tasks", c - 1, row, None)
            for key, value in {"UID": 201, "Level": 3, "Name": "Security review", "Dur": 2, "P1": 13}.items():
                calc.set("Tasks", cols[H[key]] - 1, row, value)
            calc.set("Tasks", cols[H["P1"]] - 1, row + 1, 201)      # the task below now follows the new one

        wb = self.edit(insert)
        self.assertEqual(errors_in(wb), [])
        t = Tasks(wb)
        self.assertEqual(t.get(14, "Name"), "Security review")
        self.assertEqual(t.get(14, "WBS"), "3.1.3")
        self.assertEqual(t.get(14, "PF"), dt.date(2026, 10, 26))
        self.assertEqual(t.get(15, "Name"), "Payment gateway integration")
        self.assertIsNone(t.get(15, "Link"))
        self.assertEqual(t.get(15, "PS"), dt.date(2026, 10, 27))   # follows the new task
        self.assertEqual(t.get(20, "Pred"), "15,19")                 # IDs renumbered in the MSP text
        self.assertEqual(wb["Gantt"].cell(14 + FIRST - 1, 3).value.strip(), "Security review")

    def test_delete_a_predecessor(self):
        def delete(calc, cols):
            calc.delete_row("Tasks", 14 + FIRST - 1)                # remove "Payment gateway integration" (UID 14)

        wb = self.edit(delete)
        self.assertEqual(errors_in(wb), [])
        t = Tasks(wb)
        self.assertEqual(t.get(18, "Name"), "Build complete")
        self.assertEqual(t.get(18, "Link"), "Pred 1: UID 14 not found.")
        self.assertEqual(t.get(18, "PS"), dt.date(2026, 10, 29))   # still follows its other predecessor


@unittest.skipUnless(SOFFICE, "LibreOffice not installed")
class RippleTest(Tmp):
    """A late actual finish must push every successor's forecast dates."""

    def test_slip_ripples_to_successors(self):
        t = Tasks(recalc(self.path, self.dir, lambda wb: set_input(wb, 13, "AF", dt.date(2026, 11, 4))))
        self.assertEqual(t.get(14, "FS"), dt.date(2026, 11, 5))    # next working day
        self.assertEqual(t.get(14, "FF"), dt.date(2026, 11, 12))
        self.assertEqual(t.get(19, "FS"), dt.date(2026, 11, 12))   # milestone
        self.assertEqual(t.get(21, "FS"), dt.date(2026, 11, 13))   # testing
        self.assertEqual(t.get(14, "PF"), dt.date(2026, 10, 30))   # planned untouched

    def test_status_date_moves_unstarted_work(self):
        def edit(wb):
            set_input(wb, 13, "AS", None)
            set_input(wb, 13, "Pct", None)

        t = Tasks(recalc(self.path, self.dir, edit))
        self.assertEqual(t.get(13, "FS"), dt.date(2026, 10, 26))   # first working day after 25-Oct
        self.assertEqual(t.get(13, "FF"), dt.date(2026, 11, 6))
        self.assertEqual(t.get(14, "FS"), dt.date(2026, 11, 9))


@unittest.skipUnless(SOFFICE, "LibreOffice not installed")
class BlankTemplateTest(unittest.TestCase):
    def test_blank_template(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "Project_Planner_Blank.xlsx"
            build(path, "--blank")
            wb = recalc(path, tmp, status_date=False)
            self.assertEqual(errors_in(wb), [])
            t = Tasks(wb)
            self.assertTrue(all(t.ws.cell(r, t.col[H["Name"]]).value is None for r in range(FIRST, FIRST + 200)))
            self.assertIsNone(wb["Resources"]["B6"].value)
            self.assertIsNone(wb["Holidays"]["B5"].value)
            self.assertIsNone(named_cell(wb, "ProjectStartDate").value)      # required input, highlighted yellow
            self.assertIsNone(named_cell(wb, "StatusDateInput").value)
            dash = {wb["Dashboard"].cell(r, 2).value: wb["Dashboard"].cell(r, 3).value for r in range(5, 20)}
            self.assertEqual(dash["Detail tasks"], 0)


if __name__ == "__main__":
    unittest.main()
