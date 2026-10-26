"""Agile Report: story-point KPIs, velocity forecast, the current sprint's backlog and daily burndown, and charts.

Per-sprint numbers (committed, done, velocity, release burndown) live in the Sprints sheet's SprintTable, one
formula per column; this sheet adds the project-level forecast and the current sprint, and charts both. Story
points are optional: with none entered the report shows zeros and a count of unestimated tasks.
"""
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font

from .core import DATE, Col, Show, Tab, add_table, font, name, ranked_table, title

SHEET = "Agile Report"

# defined name, label, formula, format
KPIS = [
    ("agl_Total", "Total story points (scope)", "TotalPoints", "0.#"),
    ("agl_Done", "Done by the status date", 'SUMIFS(tsk_LeafPoints,tsk_DoneDate,"<="&StatusDate)', "0.#"),
    ("agl_Remaining", "Remaining", "agl_Total-agl_Done", "0.#"),
    ("agl_Velocity", "Velocity (average of the last 3 closed sprints)",
     'IFERROR(LOOKUP(2,1/(spr_Velocity<>""),spr_Velocity),"")', "0.0"),
    ("agl_SprintsLeft", "Sprints needed at that velocity",
     'IF(N(agl_Velocity)>0,ROUNDUP(agl_Remaining/agl_Velocity,0),"")', "0"),
    ("agl_LastClosed", "Sprints closed", 'COUNTIF(spr_Status,"Closed")', "0"),
    ("agl_FinishSprint", "Forecast to finish in",
     'IF(agl_SprintsLeft="","",IF(agl_LastClosed+agl_SprintsLeft<=SprintCount,'
     'INDEX(spr_Name,agl_LastClosed+agl_SprintsLeft),"after "&INDEX(spr_Name,SprintCount)))', "@"),
    ("agl_FinishDate", "Forecast finish (velocity)",
     'IF(agl_SprintsLeft="","",IF(agl_LastClosed+agl_SprintsLeft<=SprintCount,'
     'INDEX(spr_End,agl_LastClosed+agl_SprintsLeft),""))', DATE),
    ("agl_ScheduleFinish", "Forecast finish (schedule)", "ProjectFinish", DATE),
    ("agl_Current", "Current sprint", 'IF(CurrentSprint="","(outside sprint range)",CurrentSprint)', "@"),
    ("agl_CurCommitted", "Current sprint: committed points", "SUMIFS(tsk_LeafPoints,tsk_Sprint,CurrentSprint)", "0.#"),
    ("agl_CurDone", "Current sprint: done so far",
     'SUMIFS(tsk_LeafPoints,tsk_Sprint,CurrentSprint,tsk_DoneDate,"<="&StatusDate)', "0.#"),
    ("agl_Unestimated", "Tasks without story points (not milestones)",
     'COUNTIFS(tsk_IsLeaf,1,tsk_StoryPoints,"",tsk_Duration,">0")', "0"),
]
BURNDOWN_DAYS = 31


def build_agile_report(wb, sprints_ws, sprint_letters, n_chart_sprints=12):
    ws = wb.create_sheet(SHEET)
    title(ws, None)
    ws["A1"] = '=ProjectName&"  -  Agile Report"'
    ws["A2"] = ('="Status date "&TEXT(StatusDate,"ddd dd-mmm-yyyy")&"   Sprint length "&SprintLength&'
                '" days. Per-sprint numbers: Sprints sheet. Enter Story Points (and optionally Planned Sprint) on Tasks."')
    for col, width in {"A": 3, "B": 4, "C": 44, "D": 16, "E": 30, "F": 16, "G": 12, "H": 12, "I": 12}.items():
        ws.column_dimensions[col].width = width

    top = 4
    ws.cell(top - 1, 3, "RELEASE").font = font(bold=True, color="1F4E78")
    add_table(ws, "AgileKPIs", top, 3, [
        Col("Measure", width=44, values=[k[1] for k in KPIS]),
        Col("Value", width=16, values=[f"={k[2]}" for k in KPIS]),
    ], len(KPIS), autofilter=False)
    for i, (defined, _, _, fmt) in enumerate(KPIS):
        r = top + 1 + i
        ws.cell(r, 3).font = font(bold=True)
        ws.cell(r, 4).number_format = fmt
        ws.cell(r, 4).alignment = Alignment(horizontal="right")
        name(wb, defined, f"'{SHEET}'!$D${r}")

    # current sprint backlog (hidden _Task column in B)
    top += len(KPIS) + 3
    backlog_rows = 20
    ranked_table(ws, "SprintBacklog", top, 2, "tsk_CurrentSprintKey", backlog_rows, [
        Show("Task", "tsk_Name", width=44), Show("Story Points", "tsk_StoryPoints", "num", fmt="0.#;;-"),
        Show("Resource", "tsk_Resource", width=30), Show("Forecast Finish", "tsk_FcFinish", "date", fmt=DATE),
        Show("Progress", "tsk_Progress", "num", fmt="0%"), Show("Status", "tsk_Status"),
    ], label='"Current sprint backlog ("&agl_Current&")"')

    # daily burndown of the current sprint: one row per calendar day
    top += backlog_rows + 3
    b = Tab("SprintBurndown")
    day = b.row("Day")
    k = f"ROW()-ROW({b.head('Day')})"
    ws.cell(top - 1, 3, "CURRENT SPRINT BURNDOWN (story points of the tasks in the sprint)").font = font(bold=True, color="1F4E78")
    add_table(ws, "SprintBurndown", top, 3, [
        Col("Day", width=44, fmt=DATE, formula=f'IF(OR(CurrentSprint="",{k}>SprintLength),"",CurrentSprintStart+{k}-1)'),
        Col("Remaining", width=16, fmt="0.#", formula=(
            f'IF(OR({day}="",{day}>StatusDate),"",agl_CurCommitted-SUMIFS(tsk_LeafPoints,tsk_Sprint,CurrentSprint,'
            f'tsk_DoneDate,"<="&{day}))'),
            note="Committed points not yet done at the end of the day (blank after the status date)."),
        Col("Ideal", width=30, fmt="0.#", formula=(
            f'IF({day}="","",agl_CurCommitted*MAX(0,1-NETWORKDAYS.INTL(CurrentSprintStart,{day},WorkWeek,Holidays)/'
            f'MAX(1,NETWORKDAYS.INTL(CurrentSprintStart,CurrentSprintEnd,WorkWeek,Holidays))))'),
            note="Straight line over the sprint's working days."),
        # chart copies: text labels (blank rows stay blank) and #N/A for days not reached, which charts leave out
        Col("_ChartDay", "help", 14, formula=f'IF({day}="","",TEXT({day},"dd-mmm"))'),
        Col("_ChartRemaining", "help", 14, formula=f'IF({b.row("Remaining")}="",NA(),{b.row("Remaining")})'),
        Col("_ChartIdeal", "help", 14, formula=f'IF({b.row("Ideal")}="",NA(),{b.row("Ideal")})'),
    ], BURNDOWN_DAYS, autofilter=False)
    helpers = f"F{top + 1}:H{top + BURNDOWN_DAYS}"
    ws.conditional_formatting.add(helpers, FormulaRule(formula=[f"ISNA(F{top + 1})"], font=Font(color="FFFFFF")))
    burndown_top = top

    _charts(ws, sprints_ws, sprint_letters, n_chart_sprints, burndown_top)
    ws.freeze_panes = "A4"
    return ws


def _charts(ws, sprints_ws, L, n, burndown_top):
    col = lambda key: sprints_ws[f"{L[key]}5"].column
    first, last = 5, 5 + n                      # SprintTable header row and the first n sprints (charts need A1)
    cats = Reference(sprints_ws, min_col=col("Sprint"), min_row=first + 1, max_row=last)

    release = LineChart()
    release.title, release.y_axis.title = "Release burndown (story points)", "Points"
    for key, label in (("_ChartRemaining", "Remaining"), ("Forecast Remaining", "Forecast"), ("Ideal Remaining", "Ideal")):
        release.series.append(Series(Reference(sprints_ws, min_col=col(key), min_row=first + 1, max_row=last), title=label))
    release.set_categories(cats)
    release.height, release.width = 8, 16
    ws.add_chart(release, "J4")

    velocity = BarChart()
    velocity.title = "Committed vs done points per sprint"
    for key in ("Committed Points", "Done Points"):
        velocity.add_data(Reference(sprints_ws, min_col=col(key), min_row=first, max_row=last), titles_from_data=True)
    velocity.set_categories(cats)
    velocity.height, velocity.width = 8, 16
    ws.add_chart(velocity, "J21")

    sprint = LineChart()
    sprint.title = "Current sprint burndown"
    for c, label in ((7, "Remaining"), (8, "Ideal")):       # _ChartRemaining, _ChartIdeal
        sprint.series.append(Series(Reference(ws, min_col=c, min_row=burndown_top + 1,
                                              max_row=burndown_top + BURNDOWN_DAYS), title=label))
    sprint.set_categories(Reference(ws, min_col=6, min_row=burndown_top + 1, max_row=burndown_top + BURNDOWN_DAYS))
    sprint.height, sprint.width = 8, 16
    for chart in (release, sprint):
        chart.visible_cells_only = False                    # the chart copies sit in hidden / white columns
        for s in chart.series:
            s.smooth = False
    ws.add_chart(sprint, f"J{burndown_top}")
