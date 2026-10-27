"""The Gantt timeline, drawn to the right of the Tasks table on the same rows (B-1005: one plan sheet).

The timeline cells hold no formulas at all: bars, the planned outline, milestones, deadlines, the status date and
non-working days are conditional-format rules reading the Tasks row's own helper columns (_DoneFrom, _RemFrom,
_ShowPlan, ...), and Excel only evaluates them for the cells on screen. Only the header rows (dates, month and year
labels) are formulas, so the timeline adds a few hundred cells to the calculation chain whatever the plan's size
(see docs/ARCHITECTURE.md). Planned dates are a border, the current schedule is the fill.
"""
from openpyxl.formatting.rule import FormulaRule, Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.numbers import NumberFormat
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

from .core import HEADER_FILLS, fill, font, name
from .tasks import HEADER_ROW, LAYOUT

WD = "WorkWeek,Holidays"
# planned (baseline) = border, current schedule = fill (docs/REAL_WORLD_FIT.md): red is kept for slip
COLOURS = {"critical": "F4B183", "summary": "404040", "done": "70AD47", "remaining": "9DC3E6", "slip": "E06666",
           "milestone": "7030A0", "deadline": "C00000", "planned": "000000", "planned_gap": "C9D5E8", "status": "FF0000",
           "nonworking": "EDEDED"}


def period_end(d):
    """The day after the timeline column that starts on d: GanttPeriod 1 / 7 are days, 30 / 91 calendar months /
    quarters (EDATE, so months keep their real lengths)."""
    return f"IF(GanttPeriod<28,{d}+GanttPeriod,EDATE({d},IF(GanttPeriod>=90,3,1)))"


def column_start(k):
    """The first day of timeline column k (1 = first). Month and quarter scales start on the 1st of GanttStart's
    month / calendar quarter."""
    first_month = ("DATE(YEAR(GanttStart),MONTH(GanttStart)-IF(GanttPeriod>=90,MOD(MONTH(GanttStart)-1,3),0),1)")
    return (f"IF(GanttPeriod<28,GanttStart+({k}-1)*GanttPeriod,"
            f"EDATE({first_month},({k}-1)*IF(GanttPeriod>=90,3,1)))")


def build_timeline(wb, ws, L, n_rows, n_days):
    """Header rows and conditional formatting for n_days timeline columns after the Tasks table (L: its letters).
    A narrow spacer column keeps the table from growing into the timeline when someone types next to it."""
    last_table_col = column_index_from_string(L[LAYOUT[-1][0]])
    origin_col = last_table_col + 1
    ws.column_dimensions[get_column_letter(origin_col)].width = 1
    name(wb, "gantt_Origin", f"Tasks!${get_column_letter(origin_col)}${HEADER_ROW}")
    first, last = origin_col + 1, origin_col + n_days
    k = "(COLUMN()-COLUMN(gantt_Origin))"
    for col in range(first, last + 1):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = 2.9
        day = column_start(k)
        ws.cell(HEADER_ROW, col, f"={day}").number_format = "d"
        # row 4: month (days / weeks) or year (months / quarters) where it changes; row 3: weekday, month or quarter
        ws.cell(4, col, f'=IF(GanttPeriod<28,IF(OR({k}=1,DAY({day})<=GanttPeriod),TEXT({day},"mmm yy"),""),'
                        f'IF(OR({k}=1,MONTH({day})<=IF(GanttPeriod>=90,3,1)),TEXT({day},"yyyy"),""))')
        ws.cell(3, col, f'=IF(GanttPeriod=1,LEFT(TEXT({day},"ddd"),1),IF(GanttPeriod=30,LEFT(TEXT({day},"mmm"),1),'
                        f'IF(GanttPeriod>=90,"Q"&INT((MONTH({day})+2)/3),"")))')
        for r, size, colour in ((3, 8, "404040"), (4, 8, "404040"), (5, 8, "FFFFFF")):
            c = ws.cell(r, col)
            c.font = font(size=size, bold=(r == 4), color=colour)
            c.alignment = Alignment(horizontal="left" if r == 4 else "center", vertical="center")
        ws.cell(HEADER_ROW, col).fill = HEADER_FILLS["calc"]
    _timeline_rules(ws, L, first, last, n_rows)
    _legend(ws, first)
    # a link from the top of the table to the timeline (it sits past the collapsed column groups)
    cell = f"{L['ResIn']}3"                      # clear of the next-free-UID note in A3
    jump = ws[cell]
    jump.value = "Gantt timeline >>"
    jump.hyperlink = Hyperlink(ref=cell, location=f"Tasks!{get_column_letter(first)}{HEADER_ROW}")
    jump.font = font(color="0563C1", bold=True)
    return first, last


def _timeline_rules(ws, L, first, last, n_rows):
    top, bottom = HEADER_ROW + 1, HEADER_ROW + n_rows
    g1 = get_column_letter(first)
    grid = f"{g1}{top}:{get_column_letter(last)}{bottom}"
    d = f"{g1}${HEADER_ROW}"                                   # this column's date (header row)
    at = lambda col: f"${L[col]}{top}"                         # this row's helper cell
    end = period_end(d)
    in_period = lambda x: f"AND({x}>0,{x}>={d},{x}<{end})"
    overlaps = lambda a, b: f"AND({a}>0,{d}<={b},{end}-1>={a})"
    thin = lambda colour: Side(style="thin", color=colour)
    ps, pf = at("PS"), at("PF")
    planned = f"AND({at('ShowPlan')},ISNUMBER({ps}),{overlaps(ps, pf)})"
    first_day, last_day = in_period(ps), f"AND({pf}>0,{pf}>={d},{pf}<{end})"
    edge = lambda *sides: Border(**{s: thin(COLOURS["planned"]) for s in sides})
    # Excel applies every rule that is true and, for each property, the first one listed wins
    rules = [
        FormulaRule(formula=[in_period(f"N({at('DL')})")], fill=fill(COLOURS["deadline"])),
        FormulaRule(formula=[f"AND({at('IsSum')},{overlaps(at('FS'), at('FF'))})"],
                    fill=fill(COLOURS["summary"])),
        FormulaRule(formula=[f"AND({at('IsMile')},{in_period(at('FS'))})"], fill=fill(COLOURS["milestone"])),
        FormulaRule(formula=[overlaps(at("GDoneFrom"), at("GDoneTo"))], fill=fill(COLOURS["done"])),
        # slip: remaining work after the planned finish (only where the outline is drawn)
        FormulaRule(formula=[f"AND({at('ShowPlan')},ISNUMBER({pf}),{d}>{pf},{overlaps(at('GRemFrom'), at('GRemTo'))})"],
                    fill=fill(COLOURS["slip"])),
        FormulaRule(formula=[f"AND({at('Crit')}=\"Yes\",{overlaps(at('GRemFrom'), at('GRemTo'))})"], fill=fill(COLOURS["critical"])),
        FormulaRule(formula=[overlaps(at("GRemFrom"), at("GRemTo"))], fill=fill(COLOURS["remaining"])),
        # planned days the current bar has left (work moved later or earlier): a pale gap inside the outline
        FormulaRule(formula=[planned], fill=fill(COLOURS["planned_gap"])),
        # the planned outline: top and bottom on every planned day, closed on its first and last day
        FormulaRule(formula=[f"AND({planned},{first_day},{last_day})"], border=edge("top", "bottom", "left", "right")),
        FormulaRule(formula=[f"AND({planned},{first_day})"], border=edge("top", "bottom", "left")),
        FormulaRule(formula=[f"AND({planned},{last_day})"], border=edge("top", "bottom", "right")),
        FormulaRule(formula=[planned], border=edge("top", "bottom")),
        FormulaRule(formula=[in_period("StatusDate")],
                    border=Border(left=thin(COLOURS["status"]), right=thin(COLOURS["status"]))),
        FormulaRule(formula=[f"AND(GanttPeriod=1,NETWORKDAYS.INTL({d},{d},{WD})=0)"], fill=fill(COLOURS["nonworking"])),
    ]
    for rule in rules:
        ws.conditional_formatting.add(grid, rule)
    header = f"{g1}{HEADER_ROW}:{get_column_letter(last)}{HEADER_ROW}"
    ws.conditional_formatting.add(header, FormulaRule(formula=[in_period("StatusDate")], fill=fill("C00000")))
    # month / quarter scales: rows 3-4 name the month and year, so the header's day number (always 1) is hidden
    hide = DifferentialStyle(numFmt=NumberFormat(numFmtId=200, formatCode=";;;"))
    ws.conditional_formatting.add(header, Rule(type="expression", formula=["GanttPeriod>=28"], dxf=hide))


def _legend(ws, first):
    planned = Side(style="thin", color=COLOURS["planned"])
    items = [("Planned (outline)", None, Border(top=planned, bottom=planned, left=planned, right=planned)),
             ("Done (% complete)", COLOURS["done"], None), ("Remaining - on track", COLOURS["remaining"], None),
             ("Slip: past the planned finish", COLOURS["slip"], None), ("Remaining - critical", COLOURS["critical"], None),
             ("Planned, no work now", COLOURS["planned_gap"], None), ("Summary (forecast)", COLOURS["summary"], None),
             ("Milestone (forecast)", COLOURS["milestone"], None), ("Deadline", COLOURS["deadline"], None),
             ("Status date", None, Border(left=Side(style="medium", color=COLOURS["status"]))),
             ("Non-working day", COLOURS["nonworking"], None)]
    for i, (text, colour, border) in enumerate(items):
        row, col = 1 + i // 6, first + (i % 6) * 11      # rows 1-2 only: rows 3-4 carry the month / year scale
        swatch = ws.cell(row, col)
        if colour:
            swatch.fill = fill(colour)
        if border:
            swatch.border = border
        ws.cell(row, col + 1, text).font = font(size=9)
