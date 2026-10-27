"""Schedule Report: phase summary, milestones, late tasks, the look-ahead window and open critical tasks.

Every list is a ranked_table: row k shows the task with the k-th key in a Tasks helper column (_PhaseKey,
_MilestoneOrder, _LateKey, ...). Tasks works out which tasks belong on each list and in what order; this sheet only
looks them up, so it adds one MATCH per row and no per-task SUMPRODUCT. Lists refresh with every edit; nothing needs
a button or a Power Query refresh.
"""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font

from . import rag
from .core import DATE, Col, Show, Tab, fill, ranked_table, title

FIRST_COL = 2            # column B: every list's hidden _Task column, so hiding it never hides a visible column
VAR = "+0;-0;0"
MONEY = "#,##0"
WBS = Show("WBS", "tsk_WBS", width=8)
RESOURCE = Show("Resource", "tsk_ResourceNames", width=16)
CRITICAL = Show("Critical", "tsk_Critical", width=9)
STATUS = Show("Status", "tsk_Status", width=12)


def _phase_extras():
    """SPI, CPI, open critical tasks and a RAG for each phase (from the phase's rolled-up values)."""
    p = Tab("PhaseSummary")
    blank = f'{p.row("_Task")}=""'
    spi = f'IF({blank},"",IF(N({p.row("PV")})>0,{p.row("EV")}/{p.row("PV")},""))'
    cpi = f'IF({blank},"",IF(N({p.row("AC")})>0,{p.row("EV")}/{p.row("AC")},""))'
    critical = f'IF({blank},"",COUNTIFS(tsk_Phase,{p.row("WBS")},tsk_Critical,"Yes",tsk_IsLeaf,1))'
    worst = f'MAX({rag.slip(p.row("Finish Var"))},{rag.index(p.row("SPI"))},{rag.index(p.row("CPI"))})'
    return [
        Col("SPI", width=8, fmt="0.00", formula=spi, center=True, note="Earned value / planned value. Below 1 = behind."),
        Col("CPI", width=8, fmt="0.00", formula=cpi, center=True, note="Earned value / actual cost. Below 1 = over cost."),
        Col("Open Critical", width=9, fmt="0", formula=critical, center=True, note="Unfinished critical tasks in the phase."),
        Col("RAG", width=9, formula=f'IF({blank},"",{rag.rag_text(worst)})', center=True,
            note="Worst of finish slip, SPI and CPI against the Settings thresholds."),
    ]


LISTS = [
    # table, key, rows, largest first, label, columns
    ("MilestoneReport", "tsk_MilestoneOrder", 15, False, '"Milestones (in planned order)"', [
        WBS, Show("Milestone", "tsk_Name", width=40), Show("Planned", "tsk_PlanFinish", "date", fmt=DATE),
        Show("Forecast", "tsk_FcFinish", "date", fmt=DATE), Show("Finish Var", "tsk_Variance", "num", fmt=VAR),
        Show("Deadline", "tsk_Deadline", "date", fmt=DATE), Show("Deadline Check", "tsk_DeadlineCheck"), STATUS]),
    ("LateTasks", "tsk_LateKey", 15, True, '"Late tasks (most days late first)"', [
        WBS, Show("Task", "tsk_Name", width=40), RESOURCE, Show("Planned Finish", "tsk_PlanFinish", "date", fmt=DATE),
        Show("Finish", "tsk_FcFinish", "date", fmt=DATE), Show("Finish Var", "tsk_Variance", "num", fmt=VAR),
        STATUS, CRITICAL]),
    ("StartingSoon", "tsk_StartSoonKey", 15, False,
     '"Starting by "&TEXT(LookaheadEnd,"ddd dd-mmm")&" (look-ahead "&LookaheadDays&" working days)"', [
        WBS, Show("Task", "tsk_Name", width=40), RESOURCE, Show("Start", "tsk_FcStart", "date", fmt=DATE),
        Show("Finish", "tsk_FcFinish", "date", fmt=DATE), Show("Duration", "tsk_Duration", "num"),
        Show("Predecessors", "tsk_Predecessors"), CRITICAL]),
    ("DueSoon", "tsk_DueSoonKey", 15, False, '"Due by "&TEXT(LookaheadEnd,"ddd dd-mmm")', [
        WBS, Show("Task", "tsk_Name", width=40), RESOURCE, Show("Finish", "tsk_FcFinish", "date", fmt=DATE),
        Show("Progress", "tsk_Progress", "num", fmt="0%"), Show("Deadline Check", "tsk_DeadlineCheck"), STATUS, CRITICAL]),
    ("CriticalTasks", "tsk_CriticalKey", 20, False, '"Open critical tasks (by forecast start)"', [
        WBS, Show("Task", "tsk_Name", width=40), RESOURCE, Show("Start", "tsk_FcStart", "date", fmt=DATE),
        Show("Finish", "tsk_FcFinish", "date", fmt=DATE), Show("Total Slack", "tsk_TotalSlack", "num"),
        Show("Deadline Check", "tsk_DeadlineCheck"), STATUS]),
]
PHASE_ROWS = 10


def build_schedule_report(wb):
    ws = wb.create_sheet("Schedule Report")
    title(ws, None)
    ws["A1"] = '=ProjectName&"  -  Schedule Report"'
    ws["A2"] = ('="Status date "&TEXT(StatusDate,"ddd dd-mmm-yyyy")&"   Forecast finish "&TEXT(ProjectFinish,'
                '"ddd dd-mmm-yyyy")&" ("&TEXT(FinishVarianceDays,"+0;-0;0")&" working days vs plan)"')
    ws.column_dimensions["A"].width = 3
    top = 5
    phase = ranked_table(ws, "PhaseSummary", top, FIRST_COL, "tsk_PhaseKey", PHASE_ROWS, [
        WBS, Show("Phase", "tsk_Name", width=40), Show("Planned Finish", "tsk_PlanFinish", "date", fmt=DATE),
        Show("Finish", "tsk_FcFinish", "date", fmt=DATE), Show("Finish Var", "tsk_Variance", "num", fmt=VAR),
        Show("Progress", "tsk_Progress", "num", fmt="0%"), Show("Budget", "tsk_Budget", "num", fmt=MONEY),
        Show("PV", "tsk_PV", "num", fmt=MONEY), Show("EV", "tsk_EV", "num", fmt=MONEY),
        Show("AC", "tsk_AC", "num", fmt=MONEY)], label='"Phases (outline level 1)"', extra=_phase_extras())
    last = top + PHASE_ROWS
    rag.colour(ws, f"{phase['RAG']}{top + 1}:{phase['RAG']}{last}", f"{phase['RAG']}{top + 1}")
    _late_red(ws, phase["Finish Var"], top + 1, last)
    for table, key, rows, largest, label, shows in LISTS:
        top = last + 3
        letters = ranked_table(ws, table, top, FIRST_COL, key, rows, shows, largest, label)
        last = top + rows
        if "Finish Var" in letters:
            _late_red(ws, letters["Finish Var"], top + 1, last)
        if "Critical" in letters:
            c = letters["Critical"]
            ws.conditional_formatting.add(f"{c}{top + 1}:{c}{last}", FormulaRule(
                formula=[f'{c}{top + 1}="Yes"'], font=Font(color="FFFFFF", bold=True), fill=fill("C00000")))
        if "Deadline Check" in letters:
            c = letters["Deadline Check"]
            ws.conditional_formatting.add(f"{c}{top + 1}:{c}{last}", FormulaRule(
                formula=[f'LEFT({c}{top + 1},4)="Late"'], font=Font(color="C00000", bold=True)))
    ws.page_setup.orientation = "landscape"     # one page wide when printed or saved as PDF
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.freeze_panes = "C5"
    return ws


def _late_red(ws, col, first, last):
    ws.conditional_formatting.add(f"{col}{first}:{col}{last}", FormulaRule(
        formula=[f"AND(ISNUMBER({col}{first}),{col}{first}>0)"], font=Font(color="C00000", bold=True)))
