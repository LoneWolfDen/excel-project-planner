"""Sprints sheet: sprint calendar (Excel Table "SprintTable") and resource-by-sprint load and utilisation grids.

Grid cells find their resource and sprint by position, ROW()/COLUMN() minus a named anchor cell, so every cell in
a grid holds the same formula and nothing refers to a fixed address.
"""
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from .core import (BOX, CALC_FONT, DATE, LABEL_FONT, NOTE_FONT, Col, Tab, add_table, fill, font, header_cell, name)
from .leave import leave_days, leave_row, month_column, month_start, team_leave_hours

LEAVE_MONTHS = 3        # calendar months a sprint can touch (any sprint up to about two months long)

WD = "WorkWeek,Holidays"
GOALS = ["Charter & kickoff", "Requirements", "Design & architecture", "Foundations: DB + components",
         "Core APIs + screens", "Payments + screens", "Feature complete", "System testing", "UAT",
         "Release & hypercare", "Hypercare", "Close-out"]


def work_upto(d, resource=None):
    """Hours of task work scheduled up to and including working day d (counted from the project start).

    W(D) = SUMIFS(work, finish < D) + D * SUMIFS(work per day, running on D) - SUMIFS(work before start, running on D).
    Three SUMIFS are far cheaper than a SUMPRODUCT over every task, and SUMIFS needs no array formulas."""
    res = f',tsk_Resource,{resource}' if resource else ""
    running = f'tsk_StartWD,"<="&{d},tsk_FinishWD,">="&{d}{res}'
    return (f'(SUMIFS(tsk_LeafWork,tsk_FinishWD,"<"&{d},tsk_FinishWD,">0"{res})'
            f'+{d}*SUMIFS(tsk_WorkPerWD,{running})-SUMIFS(tsk_WorkBeforeStart,{running}))')


def work_between(start, end, resource=None):
    """Hours of task work scheduled from date start to date end (inclusive), counting working days only:
    the difference of two cumulative totals (work_upto)."""
    first = f"NETWORKDAYS.INTL(ProjectStart,{start},{WD})"
    last = f"NETWORKDAYS.INTL(ProjectStart,{end},{WD})"
    return f"({work_upto(last, resource)}-{work_upto(f'({first}-1)', resource)})"


def build_sprints(wb, blank, n_sprints, n_resources):
    ws = wb.create_sheet("Sprints")
    ws["A1"] = '=ProjectName&"  -  Sprint & Capacity Planning"'
    ws["A1"].font = font(size=16, bold=True, color="1F4E78")
    ws["A2"] = ("Sprint calendar from Settings (Sprint1Start, SprintLength, SprintCount); add rows to the table for more "
                "sprints. Each task's work is spread evenly over its forecast working days, then summed per sprint.")
    ws["A2"].font = NOTE_FONT
    s = Tab("SprintTable")
    r = s.row
    has = f'{r("Sprint")}=""'
    cols = [
        Col("No", width=5, formula=f"ROW()-ROW({s.head('No')})", center=True),
        Col("Sprint", width=12, formula=f'IF({r("No")}<=SprintCount,SprintPrefix&{r("No")},"")'),
        Col("Start", width=14, fmt=DATE, formula=f"Sprint1Start+({r('No')}-1)*SprintLength"),
        Col("End", width=14, fmt=DATE, formula=f"{r('Start')}+SprintLength-1"),
        Col("Working Days", width=9, formula=f"NETWORKDAYS.INTL({r('Start')},{r('End')},{WD})", center=True),
        Col("Goal", "in", 36, values=[] if blank else GOALS),
        Col("Tasks Starting", width=9, formula=f'IF({has},"",COUNTIF(tsk_Sprint,{r("Sprint")}))', center=True),
        Col("Planned Work", width=11, fmt="#,##0.0", formula=f'IF({has},"",{work_between(r("Start"), r("End"))})',
            note="Hours of task work falling in the sprint."),
        Col("Capacity", width=11, fmt="#,##0.0",
            note="Sum of Work resources' Max Units x HoursPerDay x working days, less the leave of the people in the "
                 "grids below (Leave sheet).",
            formula=(f'IF({has},"",SUMIFS(res_MaxUnits,res_Type,"Work")*HoursPerDay*{r("Working Days")}'
                     f'-{team_leave_hours("spr_LeaveGrid", r("No"))})')),
        Col("Utilisation", width=10, fmt="0%",
            formula=f'IF(OR({has},N({r("Capacity")})=0),"",{r("Planned Work")}/{r("Capacity")})'),
        Col("Status", width=12, center=True,
            formula=f'IF({has},"",IF(StatusDate>{r("End")},"Closed",IF(StatusDate>={r("Start")},"Active","Future")))'),
        # agile: story points. Done = detail tasks finished in the sprint; remaining = scope minus everything done by
        # the sprint end (or the status date), so the columns are a release burndown
        Col("Points Capacity", "in", 9, "0.#", center=True,
            note="Optional: story points the team can take on in the sprint. Committed Points turn red above it."),
        Col("Committed Points", width=10, fmt="0.#", center=True,
            formula=f'IF({has},"",SUMIFS(tsk_LeafPoints,tsk_Sprint,{r("Sprint")}))',
            note="Story points of the tasks in this sprint (Tasks > Sprint)."),
        Col("Done Points", width=9, fmt="0.#", center=True,
            formula=(f'IF(OR({has},{r("Status")}="Future"),"",SUMIFS(tsk_LeafPoints,tsk_DoneDate,">="&{r("Start")},'
                     f'tsk_DoneDate,"<="&{r("End")}))'),
            note="Story points of tasks finished during the sprint (by Actual Finish)."),
        Col("Velocity", width=9, fmt="0.0", center=True,
            formula=(f'IF(OR({has},{r("Status")}<>"Closed"),"",AVERAGE(INDEX({s.col("Done Points")},MAX(1,{r("No")}-2))'
                     f':INDEX({s.col("Done Points")},{r("No")})))'),
            note="Average Done Points of this and the two previous sprints (closed sprints only)."),
        Col("Remaining Points", width=10, fmt="0.#", center=True,
            formula=(f'IF(OR({has},{r("Status")}="Future"),"",TotalPoints-SUMIFS(tsk_LeafPoints,tsk_DoneDate,'
                     f'"<="&MIN({r("End")},StatusDate)))'),
            note="Release burndown: story points not yet done at the sprint end (or the status date)."),
        Col("Forecast Remaining", width=10, fmt="0.#", center=True,
            formula=f'IF({has},"",TotalPoints-SUMIFS(tsk_LeafPoints,tsk_FcFinish,"<="&{r("End")}))',
            note="Points still open at the sprint end if tasks finish on their forecast dates."),
        Col("Ideal Remaining", width=10, fmt="0.#", center=True,
            formula=f'IF({has},"",TotalPoints*MAX(0,1-{r("No")}/SprintCount))',
            note="Straight line from the total scope to zero at the last sprint."),
        # charts plot "" as zero; #N/A leaves a gap, so the burndown chart reads this copy (hidden)
        Col("_ChartRemaining", "help", 8, formula=f'IF({r("Remaining Points")}="",NA(),{r("Remaining Points")})'),
    ] + [Col(f"_LeaveCol{k}", "help", 6, formula=month_column(month_start(r("Start"), k)),
             note="LeaveTable column of a month the sprint touches (for the leave look-ups).")
         for k in range(LEAVE_MONTHS)]
    L = add_table(ws, "SprintTable", 5, 1, cols, n_sprints, autofilter=False)
    for column, defined in {"Sprint": "spr_Name", "Start": "spr_Start", "End": "spr_End", "Working Days": "spr_WorkDays",
                            "Planned Work": "spr_Work", "Capacity": "spr_Capacity", "Status": "spr_Status",
                            "Committed Points": "spr_Committed", "Done Points": "spr_Done", "Velocity": "spr_Velocity",
                            "Remaining Points": "spr_Remaining"}.items():
        name(wb, defined, f"SprintTable[{column}]")
    first, last = 6, 5 + n_sprints
    cf = ws.conditional_formatting
    end = L["Ideal Remaining"]
    ws.column_dimensions.group(L["_ChartRemaining"], L[f"_LeaveCol{LEAVE_MONTHS - 1}"], hidden=True, outline_level=1)
    for k in range(LEAVE_MONTHS):
        name(wb, f"spr_LeaveCol{k}", f"SprintTable[_LeaveCol{k}]")
    cf.add(f"A{first}:{end}{last}", FormulaRule(formula=[f'${L["Sprint"]}{first}=""'], font=Font(color="D9D9D9")))
    cf.add(f"A{first}:{end}{last}", FormulaRule(formula=[f'${L["Status"]}{first}="Active"'], fill=fill("FFF2CC"), font=Font(bold=True)))
    cf.add(f"{L['Utilisation']}{first}:{L['Utilisation']}{last}",
           FormulaRule(formula=[f"AND(ISNUMBER(${L['Utilisation']}{first}),${L['Utilisation']}{first}>1)"],
                       font=Font(color="C00000", bold=True), fill=fill("FFC7CE")))
    pc, cp = L["Points Capacity"], L["Committed Points"]
    cf.add(f"{cp}{first}:{cp}{last}", FormulaRule(
        formula=[f"AND(N(${pc}{first})>0,N(${cp}{first})>${pc}{first})"], font=Font(color="C00000", bold=True),
        fill=fill("FFC7CE")))

    load_top = last + 3
    util_top = load_top + n_resources + 5
    _grid(wb, ws, "load", load_top, n_sprints, n_resources,
          "Resource load by sprint (hours)  -  red = over capacity (Max Units x HoursPerDay x working days, less leave)",
          util_top - load_top)
    _grid(wb, ws, "util", util_top, n_sprints, n_resources, "Resource utilisation by sprint (% of capacity after leave)")
    _leave_grid(wb, ws, util_top + n_resources + 5, n_sprints, n_resources)
    ws.freeze_panes = "D6"
    ws.column_dimensions["B"].width = 16
    return ws, L


def _grid(wb, ws, kind, top, n_sprints, n_resources, heading, util_offset=0):
    """Rows = resources, columns = sprints. The origin cell (resource header) anchors every position."""
    ws.cell(top - 1, 1, heading).font = font(bold=True, size=12, color="1F4E78")
    origin = f"{kind}_Origin"
    name(wb, origin, f"Sprints!$B${top}")
    header_cell(ws.cell(top, 2), "Resource")
    header_cell(ws.cell(top, 3), "Max Units")
    i = f"(ROW()-ROW({origin})-1)"                 # resource number (row 1 below the dates row)
    j = f"(COLUMN()-COLUMN({origin})-1)"            # sprint number
    sprint_name = f"INDEX(spr_Name,{j})"
    for c in range(4, 4 + n_sprints):
        header_cell(ws.cell(top, c), None)
        ws.cell(top, c).value = f'=IFERROR({sprint_name}&"","")'
        d = ws.cell(top + 1, c, f'=IFERROR(IF({sprint_name}="","",INDEX(spr_Start,{j})),"")')
        d.number_format = "dd-mmm"
        d.font = font(italic=True, size=9)
        ws.column_dimensions[get_column_letter(c)].width = 10
    ws.cell(top + 1, 2, "Sprint start").font = font(italic=True, size=9)
    resource = f"INDEX(res_Name,{i})"
    units = f"INDEX(res_MaxUnits,{i})"
    first_row, last_row = top + 2, top + 1 + n_resources
    last_col = get_column_letter(3 + n_sprints)
    for rr in range(first_row, last_row + 1):
        ws.cell(rr, 2, f'=IFERROR({resource}&"","")').font = font(bold=True)
        u = ws.cell(rr, 3, f'=IFERROR(IF({resource}="","",{units}),"")')
        u.number_format = "0%"
        for c in range(4, 4 + n_sprints):
            if kind == "load":
                hours = work_between(f"INDEX(spr_Start,{j})", f"INDEX(spr_End,{j})", resource)
                v = f'=IFERROR(IF(OR({resource}="",{sprint_name}=""),"",{hours}),"")'
                fmt = "#,##0;;-"
            else:
                load = f"INDEX(load_Grid,{i},{j})"
                capacity = f"({units}*HoursPerDay*(INDEX(spr_WorkDays,{j})-INDEX(spr_LeaveGrid,{i},{j})))"
                v = (f'=IFERROR(IF(OR({resource}="",{sprint_name}="",N({units})=0),"",'
                     f'IF({capacity}<=0,IF(N({load})>0,9.99,""),{load}/{capacity})),"")')
                fmt = "0%;;-"
            cell = ws.cell(rr, c, v)
            cell.number_format = fmt
            cell.border = BOX
            cell.font = CALC_FONT
    grid = f"D{first_row}:{last_col}{last_row}"
    if kind == "load":
        name(wb, "load_Grid", f"Sprints!$D${first_row}:${last_col}${last_row}")
        util_cell = f"D{first_row + util_offset}"     # the matching utilisation cell: leave counts
        ws.conditional_formatting.add(grid, FormulaRule(
            formula=[f"AND(ISNUMBER({util_cell}),{util_cell}>1)"], fill=fill("FFC7CE"), font=Font(color="9C0006", bold=True)))
    else:
        ws.conditional_formatting.add(grid, ColorScaleRule(
            start_type="num", start_value=0, start_color="FFFFFF", mid_type="num", mid_value=0.8, mid_color="C6EFCE",
            end_type="num", end_value=1.2, end_color="F8696B"))


def _leave_grid(wb, ws, top, n_sprints, n_resources):
    """Hidden grid: working days of leave per person per sprint (Leave sheet), read by the utilisation grid. Each
    person's LeaveTable row is looked up once, in column C; each sprint's month columns come from SprintTable."""
    origin = "spr_LeaveOrigin"
    name(wb, origin, f"Sprints!$B${top}")
    ws.cell(top, 2, "Leave days by sprint (helper for the utilisation grid)").font = font(italic=True)
    i = f"(ROW()-ROW({origin}))"
    j = f"(COLUMN()-COLUMN({origin})-1)"
    resource = f"INDEX(res_Name,{i})"
    start, end = f"INDEX(spr_Start,{j})", f"INDEX(spr_End,{j})"
    months = [(month_start(start, k), f"INDEX(spr_LeaveCol{k},{j})") for k in range(LEAVE_MONTHS)]
    days = leave_days(start, end, f"INDEX(spr_LeaveRow,{i})", months)
    last_col = get_column_letter(3 + n_sprints)
    for r in range(top + 1, top + 1 + n_resources):
        ws.cell(r, 3, f'=IFERROR(IF({resource}="",0,{leave_row(resource)}),0)').font = CALC_FONT
        for c in range(4, 4 + n_sprints):
            ws.cell(r, c, f'=IFERROR(IF(OR({resource}="",INDEX(spr_Name,{j})=""),0,{days}),0)').font = CALC_FONT
    name(wb, "spr_LeaveRow", f"Sprints!$C${top + 1}:$C${top + n_resources}")
    name(wb, "spr_LeaveGrid", f"Sprints!$D${top + 1}:${last_col}${top + n_resources}")
    for r in range(top, top + 1 + n_resources):
        ws.row_dimensions[r].hidden = True
