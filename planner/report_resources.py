"""Resource Report: a summary per resource, and weekly hours and utilisation grids with over-allocation in red.

The grids work like the Sprints load grid: every cell finds its resource and week by ROW()/COLUMN() minus a named
anchor, so each grid holds one formula. Hours come from a hidden grid of cumulative work at each week boundary
(sprints.work_upto: SUMIFS over the time-phased helper columns on Tasks, working days only); a week's hours are
the difference of two neighbours. Working each boundary out once halves the SUMIFS, and boundaries before the
project or after it finishes need none. Weeks start on Settings > ResourceWeekStart, so the window scrolls by
changing one date. The summary reads totals from ResourceTable and the grids; it adds no per-task SUMPRODUCT.
"""
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from .core import BOX, CALC_FILL, CALC_FONT, Col, Tab, add_table, fill, font, header_cell, name, title
from .sprints import work_upto

SHEET = "Resource Report"
N_WEEKS = 26
WD = "WorkWeek,Holidays"


def build_resource_report(wb, n_resources):
    ws = wb.create_sheet(SHEET)
    title(ws, None)
    ws["A1"] = '=ProjectName&"  -  Resource Report"'
    ws["A2"] = ('="Weeks from "&TEXT(ResourceWeekStart,"ddd dd-mmm-yyyy")&" (Settings > ResourceWeekStart). Hours are '
                'each task\'s work spread over its forecast working days; red = more than Max Units x HoursPerDay x '
                'working days in the week."')
    ws.column_dimensions["A"].width = 3

    summary_top = 5
    _summary(ws, summary_top, n_resources)
    load_top = summary_top + n_resources + 4
    util_top = load_top + n_resources + 7
    _grid(wb, ws, "load", load_top, n_resources, "Weekly hours by resource")
    _grid(wb, ws, "util", util_top, n_resources, "Weekly utilisation by resource (hours / capacity)")
    _cumulative(wb, ws, util_top + n_resources + 4, n_resources)
    _team_chart(ws, load_top, n_resources)
    ws.freeze_panes = "D4"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def _summary(ws, top, n_resources):
    t = Tab("ResourceSummary")
    k = f"ROW()-ROW({t.head('Resource')})"
    res = t.row("Resource")
    blank = f'{res}=""'
    from_table = lambda col: f'IF({blank},"",INDEX(ResourceTable[{col}],{k}))'
    week_row = lambda grid: f"INDEX({grid},{k},0)"
    ws.cell(top - 1, 2, "RESOURCES").font = font(bold=True, color="1F4E78")
    add_table(ws, "ResourceSummary", top, 2, [
        Col("Resource", width=22, formula=f'IFERROR(INDEX(res_Name,{k})&"","")'),
        Col("Max Units", width=10, fmt="0%", formula=from_table("Max Units"), center=True),
        Col("Tasks", width=10, fmt="0", formula=from_table("Tasks"), center=True),
        Col("Work", width=10, fmt="#,##0", formula=from_table("Work"), note="Assigned hours."),
        Col("Work Done", width=10, fmt="#,##0", formula=f'IF({blank},"",SUMIFS(tsk_LeafWorkDone,tsk_Resource,{res}))',
            note="Work x Progress."),
        Col("Remaining", width=10, fmt="#,##0", formula=f'IF({blank},"",{t.row("Work")}-{t.row("Work Done")})'),
        Col("Cost", width=10, fmt="#,##0", formula=from_table("Cost")),
        Col("Actual Cost", width=10, fmt="#,##0", formula=from_table("Actual Cost")),
        Col("Late Tasks", width=10, fmt="0", center=True,
            formula=f'IF({blank},"",COUNTIFS(tsk_Resource,{res},tsk_LateKey,">0"))',
            note="Unfinished tasks forecast to finish after plan, or not started on time (Schedule Report)."),
        Col("Open Critical", width=10, fmt="0", center=True,
            formula=f'IF({blank},"",COUNTIFS(tsk_Resource,{res},tsk_Critical,"Yes",tsk_IsLeaf,1))'),
        Col("Peak Week (h)", width=10, fmt="#,##0", center=True, formula=f'IF({blank},"",MAX({week_row("rw_Load")}))',
            note="Most hours in any week shown in the grid below."),
        Col("Peak Use", width=10, fmt="0%", center=True, formula=f'IF({blank},"",MAX({week_row("rw_Util")}))',
            note="Highest weekly utilisation in the grid."),
        Col("Weeks Over", width=10, fmt="0", center=True,
            formula=f'IF({blank},"",COUNTIF({week_row("rw_Util")},">1"))',
            note="Weeks in the grid above 100% of capacity."),
    ], n_resources, autofilter=False)
    first, last = top + 1, top + n_resources
    over = f"N{first}:N{last}"
    ws.conditional_formatting.add(over, FormulaRule(formula=[f"N(N{first})>0"], font=Font(color="9C0006", bold=True),
                                                    fill=fill("FFC7CE")))
    ws.conditional_formatting.add(f"M{first}:M{last}", FormulaRule(
        formula=[f"AND(ISNUMBER(M{first}),M{first}>1)"], font=Font(color="9C0006", bold=True)))


def _grid(wb, ws, kind, top, n_resources, heading):
    """Rows = resources, columns = weeks. The origin cell (the "Resource" header) anchors every position."""
    ws.cell(top - 2, 2, heading).font = font(bold=True, size=12, color="1F4E78")
    origin = f"rw_{kind.capitalize()}Origin"
    name(wb, origin, f"'{SHEET}'!$B${top}")
    header_cell(ws.cell(top, 2), "Resource")
    header_cell(ws.cell(top, 3), "Max Units")
    i = f"(ROW()-ROW({origin})-1)"               # resource number (row 1 below the working-days row)
    j = f"(COLUMN()-COLUMN({origin})-1)"          # week number
    week = f"(ResourceWeekStart+7*({j}-1))"
    first_col, last_col = 4, 3 + N_WEEKS
    for c in range(first_col, last_col + 1):
        header_cell(ws.cell(top, c), None)
        ws.cell(top, c).value = f"={week}"
        ws.cell(top, c).number_format = "dd-mmm"
        days = ws.cell(top + 1, c, f"=NETWORKDAYS.INTL({week},{week}+6,{WD})")
        days.font = font(italic=True, size=9)
        days.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(c)].width = 9
    ws.cell(top + 1, 2, "Working days").font = font(italic=True, size=9)
    L = get_column_letter
    if kind == "load":
        name(wb, "rw_WeekDays", f"'{SHEET}'!${L(first_col)}${top + 1}:${L(last_col)}${top + 1}")
    resource, units = f"INDEX(res_Name,{i})", f"INDEX(res_MaxUnits,{i})"
    first_row, last_row = top + 2, top + 1 + n_resources
    for r in range(first_row, last_row + 1):
        ws.cell(r, 2, f'=IFERROR({resource}&"","")').font = font(bold=True)
        u = ws.cell(r, 3, f'=IFERROR(IF({resource}="","",{units}),"")')
        u.number_format = "0%"
        for c in range(first_col, last_col + 1):
            if kind == "load":
                v = f'=IFERROR(IF({resource}="","",INDEX(rw_Cum,{i},{j}+1)-INDEX(rw_Cum,{i},{j})),"")'
                fmt = "#,##0;;-"
            else:
                capacity = f"{units}*HoursPerDay*INDEX(rw_WeekDays,{j})"
                v = (f'=IFERROR(IF(OR({resource}="",N({units})=0,INDEX(rw_WeekDays,{j})=0),"",'
                     f'INDEX(rw_Load,{i},{j})/({capacity})),"")')
                fmt = "0%;;-"
            cell = ws.cell(r, c, v)
            cell.number_format, cell.border, cell.font = fmt, BOX, CALC_FONT
    grid = f"{L(first_col)}{first_row}:{L(last_col)}{last_row}"
    name(wb, f"rw_{kind.capitalize()}", f"'{SHEET}'!${L(first_col)}${first_row}:${L(last_col)}${last_row}")
    tl = f"{L(first_col)}{first_row}"
    if kind == "load":
        ws.conditional_formatting.add(grid, FormulaRule(
            formula=[f"AND(ISNUMBER({tl}),{tl}>$C{first_row}*HoursPerDay*{L(first_col)}${top + 1})"],
            fill=fill("FFC7CE"), font=Font(color="9C0006", bold=True)))
        # team totals: the chart below plots them
        totals = [("Team hours", f"SUM(INDEX(rw_Load,0,{j}))"),
                  ("Team capacity", f'SUMIFS(res_MaxUnits,res_Type,"Work")*HoursPerDay*INDEX(rw_WeekDays,{j})')]
        for n, (label, formula) in enumerate(totals):
            r = last_row + 1 + n
            ws.cell(r, 2, label).font = font(bold=True)
            for c in range(first_col, last_col + 1):
                cell = ws.cell(r, c, f"={formula}")
                cell.number_format, cell.border, cell.font, cell.fill = "#,##0;;-", BOX, font(bold=True), CALC_FILL
    else:
        ws.conditional_formatting.add(grid, ColorScaleRule(
            start_type="num", start_value=0, start_color="FFFFFF", mid_type="num", mid_value=0.8, mid_color="C6EFCE",
            end_type="num", end_value=1.2, end_color="F8696B"))


def _team_chart(ws, load_top, n_resources):
    totals_row = load_top + 2 + n_resources
    chart = BarChart()
    chart.title, chart.y_axis.title = "Team hours per week vs capacity", "Hours"
    weeks = Reference(ws, min_col=4, max_col=3 + N_WEEKS, min_row=load_top)
    chart.series.append(Series(Reference(ws, min_col=4, max_col=3 + N_WEEKS, min_row=totals_row), title="Team hours"))
    chart.set_categories(weeks)
    capacity = LineChart()
    capacity.series.append(Series(Reference(ws, min_col=4, max_col=3 + N_WEEKS, min_row=totals_row + 1),
                                  title="Team capacity"))
    capacity.set_categories(weeks)
    capacity.series[0].smooth = False
    chart.gapWidth = 40
    chart += capacity
    chart.height, chart.width = 8, 22
    ws.add_chart(chart, "P4")


def _cumulative(wb, ws, top, n_resources):
    """Hidden grid: each resource's cumulative work up to the end of the day before week b (b = 0 .. N_WEEKS).

    Week j's hours are column j+1 minus column j. Before the project starts the total is 0; from the forecast
    finish on it is the resource's whole work (one SUMIFS); only boundaries inside the project need work_upto."""
    origin = "rw_CumOrigin"
    name(wb, origin, f"'{SHEET}'!$B${top}")
    ws.cell(top, 2, "Cumulative hours at each week boundary (helper for the weekly grids)").font = font(italic=True)
    i = f"(ROW()-ROW({origin}))"
    b = f"(COLUMN()-COLUMN({origin})-2)"
    day = f"(ResourceWeekStart+7*{b}-1)"
    resource = f"INDEX(res_Name,{i})"
    total = f"SUMIFS(tsk_LeafWork,tsk_Resource,{resource})"
    upto = work_upto(f"NETWORKDAYS.INTL(ProjectStart,{day},{WD})", resource)
    formula = (f'=IFERROR(IF({resource}="","",IF({day}<ProjectStart,0,IF({day}>=ProjectFinish,{total},{upto}))),"")')
    first_col, last_col = 4, 4 + N_WEEKS
    for r in range(top + 1, top + 1 + n_resources):
        for c in range(first_col, last_col + 1):
            ws.cell(r, c, formula).font = CALC_FONT
    L = get_column_letter
    name(wb, "rw_Cum", f"'{SHEET}'!${L(first_col)}${top + 1}:${L(last_col)}${top + n_resources}")
    for r in range(top, top + 1 + n_resources):
        ws.row_dimensions[r].hidden = True
