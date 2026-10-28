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
            "Total Slack", "Work", "Status", "Critical", "Progress"]
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
        self.assertIn("3 tasks had a typed Actual Cost", warnings[0])     # cost is now work done x day rate
        renumbered = warnings[1:]
        self.assertEqual([w.split(" refs")[0] for w in renumbered], ["Decision", "Change"])   # renumbered, and said so
        self.assertIn("Decision 2 is now Decision 8", renumbered[0])

    def test_pm_rag_stays_with_its_area(self):
        wb = openpyxl.load_workbook(FIXTURES / "planner_v1.6.xlsx")
        ws = wb["Status Report"]
        ref = ws.tables["StatusRAG"].ref
        rows = list(ws[ref])
        head = [c.value for c in rows[0]]
        for row in rows[1:]:
            area = row[head.index("Area")].value
            if area in ("Cost performance", "Deadlines"):
                row[head.index("PM RAG")].value = "Amber" if area == "Deadlines" else "Green"
        data = io.BytesIO()
        wb.save(data)
        _, copies, _, warnings, _ = upgrade.read_planner(data.getvalue())
        head, rows = copies["StatusRAG"]
        pm = {r[head.index("Area")]: r[head.index("PM RAG")] for r in rows}
        self.assertNotIn("Cost performance", pm)                        # the row is gone, so nothing shifts
        self.assertEqual(pm["Deadlines"], "Amber")
        self.assertIn("Cost performance RAG is now on the Cost sheet", " ".join(warnings))

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


def _typed(data):
    """{(table, header): {typed value}} of a workbook's tables, formulas left out (what a user typed)."""
    out = {}
    for name, (_, head, rows) in upgrade._tables(openpyxl.load_workbook(io.BytesIO(data))).items():
        for i, h in enumerate(head):
            values = {r[i] for r in rows if r[i] not in (None, "")}
            if values:
                out[(name, h)] = values
    return out


def _text(v):
    return f"{v:g}" if isinstance(v, (int, float)) and not isinstance(v, bool) else str(v)


# Typed columns that a version moved or turned into something else: (table, header) -> (new table, new header,
# value -> what to find in that column). The value must appear in (or as) one of the new column's typed cells.
PERCENT = lambda v: f"[{v:.0%}]"
MOVED = {
    **{("Tasks", f"Pred{k} {x}"): ("Tasks", "Predecessors", _text) for k in range(1, 5) for x in ("UID", "Lag")},
    # FS is the default, so "12" in the text means 12FS
    **{("Tasks", f"Pred{k} Type"): ("Tasks", "Predecessors", lambda v: "" if v == "FS" else v) for k in range(1, 5)},
    ("Tasks", "Resource"): ("Tasks", "Resources", _text), ("Tasks", "Units"): ("Tasks", "Resources", PERCENT),
    ("Tasks", "Actual Cost"): ("Tasks", "Notes", _text),          # calculated since v1.8; the typed value kept in Notes
    ("Tasks", "% Complete"): ("Tasks", "Stage", lambda v: "Complete" if v >= 1 else ""),   # a Stage when it matches one
    ("AssignmentTable", "Task UID"): ("Tasks", "UID", _text),
    ("AssignmentTable", "Resource"): ("Tasks", "Resources", _text),
    ("AssignmentTable", "Units"): ("Tasks", "Resources", PERCENT),
    ("LeaveTable", "Resource"): ("ResourceTable", "Resource Name", _text),
    ("ResourceTable", "Max Units"): ("ResourceTable", "Allocation %", _text),
    ("ResourceTable", "Std Rate"): ("RateCard", "Day Rate", lambda v: _text(v * 8)),     # hourly x HoursPerDay 8
    ("StatusHistory", "% Earned"): ("StatusHistory", "% Done", _text),
    **{("StatusHistory", h): ("CostHistory", h, _text) for h in ("BAC", "PV", "EV", "AC", "CPI", "EAC")},
}
# the template's own text (row labels of report tables): rebuilt by the new template, never user input
TEMPLATE = {("AgileKPIs", "Measure"), ("ChangeControl", "Measure"), ("ChangeControl", "Source"),
            ("EarnedValue", "Measure"), ("EarnedValue", "Meaning"), ("KPIs", "Key indicator"),
            ("StatusRAG", "Area"), ("StatusRAG", "Measure")}


class TypedDataTest(unittest.TestCase):
    """B-1012: every value typed into an old workbook reaches the upgraded one, under its own table and header, a
    renamed one (upgrade.RENAMED, RENAMED_TABLES, LOGS) or a listed move (MOVED). A typed column nobody mapped
    fails here, so a new version can't drop it silently."""

    def check(self, fixture):
        data = (FIXTURES / fixture).read_bytes()
        old = _typed(data)
        xlsx, _ = upgrade.convert(data)
        new = _typed(xlsx)
        lost = []
        for (table, header), values in sorted(old.items()):
            if (table, header) in TEMPLATE:
                continue
            if (table, header) in MOVED:
                where, column, expect = MOVED[(table, header)]
            elif table in upgrade.LOGS:
                where, column, expect = "Log", upgrade.LOGS[table][1].get(header, header), _text
            elif table == "LeaveTable":
                where, column, expect = "ResourceTable", header, _text      # leave months moved onto Resources
            else:
                where, renamed = upgrade.RENAMED_TABLES.get(table, (table, {}))
                column, expect = renamed.get(header, upgrade.RENAMED.get(table, {}).get(header, header)), _text
            got = [_text(v) for v in new.get((where, column), ())]
            for v in values:
                want = expect(v)
                if want and not any(want == g or want in g for g in got):
                    lost.append(f"{table}[{header}] {v!r} -> {where}[{column}]")
        self.assertEqual(lost, [])

    def test_v1_4(self):
        self.check("planner_v1.4.xlsx")

    def test_v1_6(self):
        self.check("planner_v1.6.xlsx")


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
        # each old history column lands under its new name, or on the Cost sheet's CostHistory for the money
        history, costs = (log(new, s, t, "Status Date") for s, t in (("History", "StatusHistory"), ("Cost", "CostHistory")))
        old_history = log(old, "History", "StatusHistory", "Status Date")
        self.assertEqual((len(history), len(costs)), (len(old_history), len(old_history)))
        for o, h, c in zip(old_history, history, costs):
            for k, v in o.items():
                k = "% Done" if k == "% Earned" else k
                if k in h or k in c:
                    self.assertEqual(h[k] if k in h else c[k], v, k)
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
        # Cost performance moved to the Cost sheet's own RAG; the fixtures have no PM override on it
        self.assertEqual(rag(new, "Status"), [r for r in rag(old, "Status Report") if r[0] != "Cost performance"])
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
