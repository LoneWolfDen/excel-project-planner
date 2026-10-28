"""Leave: each person's leave, one column per month on their Resources row (ResourceTable), and the formula that reads
it.

A cell under "Nov-26" holding "16, 17(0.5)" means a full day off on 16-Nov-26 and half a day on 17-Nov-26
(semicolons work too; spaces are ignored). Leave reduces capacity only: utilisation, over-allocation and sprint
capacity count it, task dates don't move (the plan stays the same in MS Project).

Reading the text needs no helper cells:
  * the month's column is found by its header text, so columns can be added (insert one among the months and type
    "Jan-29") or moved. The label is
    spelled out with MID, because TEXT(date,"mmm") depends on Excel's language. Reports look each month's column up
    once per grid column, and each person's row once per grid row, in helper cells;
  * each day of the month is tested against the text in one SUMPRODUCT over the constant {1,...,31}
    (the name leave_Days). Excel 2019 does not evaluate IF or IFERROR on arrays inside SUMPRODUCT, so the value is
    read without them: every day in the text gets a value in brackets ("16" becomes "16(1)"), and ",N(0)" is
    appended for each day N, so the search for ",N(" always succeeds and a day not on leave reads 0. Values have
    at most four characters (0.5, 0.25, 1);
  * only working days count (WorkWeek and Holidays), so leave typed on a weekend or holiday changes nothing.
"""
from openpyxl.worksheet.datavalidation import DataValidation

from .core import Col, Tab, add_table, col_range, font, name, release

TABLE = "ResourceTable"           # the months are columns of the Resources table, after each person's details
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
DEFAULT_FIRST, DEFAULT_COUNT = (2026, 1), 36       # Jan-26 .. Dec-28; more columns can be added by typing a header
SAMPLE = {
    "Casey Kim": {"Nov-26": "16, 17(0.5), 18, 19, 20, 23, 24, 25, 26, 27"},
    "Jordan Lee": {"Oct-26": "30", "Dec-26": "21, 22, 23(0.5)"},
    "Sam Rivera": {"Nov-26": "2(0.5)"},
    "Morgan Patel": {"Dec-26": "28, 29, 30"},
}


def month_labels(first=DEFAULT_FIRST, count=DEFAULT_COUNT):
    year, month = first
    out = []
    for _ in range(count):
        out.append(f"{MONTHS[month - 1]}-{year % 100:02d}")
        year, month = (year + 1, 1) if month == 12 else (year, month + 1)
    return out


def month_start(day, offset=0):
    """The first day of the month `offset` months after day's month."""
    return f"DATE(YEAR({day}),MONTH({day})+{offset},1)"


def month_column(first_of_month):
    """Position of that month's column in LeaveTable, found by its header (e.g. Nov-26); 0 when there is none.
    The label is spelled out with MID, because TEXT(date,"mmm") depends on Excel's language."""
    label = f'MID("{"".join(MONTHS)}",MONTH({first_of_month})*3-2,3)&"-"&RIGHT(YEAR({first_of_month}),2)'
    return f"IFERROR(MATCH({label},leave_Months,0),0)"


def leave_row(resource):
    """Position of a person's row in LeaveTable; 0 when they have none."""
    return f"IFERROR(MATCH({resource},{TABLE}[Resource Name],0),0)"


def _in_month(start, end, first, text, weights=None):
    """Working days of leave in [start, end] within the month starting on `first`, from text made by _text (one
    cell, or a column of cells weighted by `weights`). One SUMPRODUCT over leave_Days = {1..31}: for day N, the
    value in brackets after ",N(" (the ",N(0)" appended for every N means the search always succeeds)."""
    n = "leave_Days"
    day = f"({first}+{n}-1)"
    padded = f'{text}&","&{n}&"(0)"'
    value = f'MID({padded},SEARCH(","&{n}&"(",{padded})+LEN({n})+2,5)'
    amount = f'--LEFT(SUBSTITUTE({value},")",REPT(" ",5)),5)'
    in_range = (f"({n}<=DAY(EOMONTH({first},0)))*({day}>={start})*({day}<={end})"
                f'*(MID(WorkWeek,WEEKDAY({day},2),1)="0")*(COUNTIF(Holidays,{day})=0)')
    weight = f"{weights}*" if weights else ""
    return f"IFERROR(SUMPRODUCT({weight}{in_range}*{amount}),0)"


def _text(cells):
    """Leave text in a form one search can read: "16, 17(0.5)" becomes "(1),16(1),17(0.5)(1)," (spaces removed,
    semicolons read as commas, and every day given a value in brackets: 1 unless one was typed)."""
    clean = f'SUBSTITUTE(SUBSTITUTE({cells}&""," ",""),";",",")'
    return f'SUBSTITUTE(","&{clean}&",",",","(1),")'


def leave_days(start, end, row, columns):
    """One person's working days of leave from start to end. row = leave_row(...); columns = [(first day of month,
    month_column(...)), ...] for each calendar month the period can touch. Pass row and columns as references to
    helper cells, so each lookup is done once per grid row / column rather than in every cell."""
    cell = lambda col: f"INDEX(leave_Grid,{row},{col})"
    if release():                                      # Release 1.0: plus the days ticked on the leave day grid
        cell = lambda col: (f"INDEX(leave_Grid,{row},{col})&\",\"&IFERROR(INDEX(lv_Days,MATCH(INDEX({TABLE}"
                            f"[Resource Name],{row})&\"|\"&INDEX(leave_Months,{col}),lv_Key,0)),\"\")")
    terms = [f'IF(OR({row}=0,{col}=0),0,{_in_month(start, end, first, _text(cell(col)))})' for first, col in columns]
    return "(" + "+".join(terms) + ")"


def team_leave_hours(grid, column):
    """Capacity hours the team's leave removes in one period: the sum over a per-person leave grid's column (people
    in ResourceTable order) of days x Max Units x HoursPerDay, for Work resources. Reading the grid, which already
    holds each person's days, is far cheaper than parsing everyone's leave text again."""
    people = f"ROWS({grid})"
    units = f"INDEX(res_Allocation,1):INDEX(res_Allocation,{people})"
    work = f'(INDEX(res_Type,1):INDEX(res_Type,{people})="Work")'
    return f"(HoursPerDay*SUMPRODUCT({work}*{units}*INDEX({grid},0,{column})))"


def leave_columns(people, blank, entries=None, months=None):
    """The month columns of ResourceTable: one per month header, holding each person's days off (people = the
    Resources rows, name first). entries: {person: {month header: days text}} replacing the sample (an import)."""
    labels = months or month_labels()
    data = entries if entries is not None else {} if blank else SAMPLE
    if release() and entries is None:                  # the Release 1.0 sample enters some leave on the day grid
        data = {p: m for p, m in data.items() if p not in GRID_SAMPLE}
    names = [p[0] for p in people]
    return [Col(label, "in", 11, values=[data.get(n, {}).get(label) for n in names],
                note=f'Days off in {label}: "16, 17(0.5)" = 16th all day, 17th half a day.') for label in labels]


LEAVE_DAYS = "{" + ",".join(str(d) for d in range(1, 32)) + "}"     # the name leave_Days: days of a month


def name_leave(wb, months):
    """leave_Grid / leave_Months: only the month columns of ResourceTable. The leave look-ups read these, never the
    whole table, so a change to a calculated column there (Work, Cost) doesn't recalculate every leave grid."""
    span = f"[{months[0]}]:[{months[-1]}]"
    name(wb, "leave_Grid", f"{TABLE}[[#Data],{span}]")
    name(wb, "leave_Months", f"{TABLE}[[#Headers],{span}]")
    name(wb, "leave_Days", LEAVE_DAYS)


GRID_ROWS = 40                    # rows of the leave day grid (Release 1.0): one per person per month with leave
GRID = "LeaveDays"
GRID_SAMPLE = {"Casey Kim": {"Nov-26": "16, 17(0.5), 18, 19, 20, 23, 24, 25, 26, 27"},
               "Jordan Lee": {"Oct-26": "30", "Dec-26": "21, 22, 23(0.5)"}}


def _days_of(text):
    """{day: amount} from leave text: "16, 17(0.5)" -> {16: 1, 17: 0.5}."""
    out = {}
    for part in str(text).replace(";", ",").split(","):
        part = part.strip()
        if part:
            day, _, amount = part.partition("(")
            out[int(day)] = float(amount.rstrip(")")) if amount else 1
    return out


def build_leave_grid(wb, ws, top, blank, n_rows=GRID_ROWS):
    """Release 1.0: the leave day grid (table LeaveDays) under the Resources table. One row per person per month:
    pick the person and the month, then type 1 (a day) or 0.5 (half a day) under each day off. Days builds the
    same text a month cell holds ("16, 17(0.5)") with TEXTJOIN, and every leave look-up adds it to the person's
    month cell, so either way of entering leave counts (once each: don't enter the same day in both).
    Returns the row under the grid."""
    ws.cell(top - 2, 2, "LEAVE BY DAY: one row per person and month; type 1 (all day) or 0.5 (half a day) under "
                        "each day off. It adds to the month columns above.").font = font(bold=True, color="1F4E78")
    data = {} if blank else GRID_SAMPLE
    entries = [(person, month, _days_of(text)) for person, months in data.items() for month, text in months.items()]
    t = Tab(GRID)
    day = lambda k: t.row(str(k))
    text = ",".join(f'IF({day(k)}="","",{k}&IF(N({day(k)})=1,"","("&N({day(k)})&")"))' for k in range(1, 32))
    cols = [Col("Resource", "in", 22, values=[e[0] for e in entries], note="Pick the person (Resources)."),
            Col("Month", "in", 9, values=[e[1] for e in entries], note="Pick the month: one of the month columns "
                                                                         "above, e.g. Nov-26."),
            *[Col(str(k), "in", 3.6, center=True, values=[e[2].get(k) for e in entries]) for k in range(1, 32)],
            Col("Days", "calc", 26, formula=f'_xlfn.TEXTJOIN(", ",TRUE(),{text})',
                note="The days off as the month cells write them: what the reports read."),
            Col("_Key", "help", 6, formula=f'{t.row("Resource")}&"|"&{t.row("Month")}',
                note="Person|month: the look-up key (helper).")]
    letters = add_table(ws, GRID, top, 2, cols, n_rows)
    ws.column_dimensions[letters["_Key"]].hidden = True
    name(wb, "lv_Days", f"{GRID}[Days]")
    name(wb, "lv_Key", f"{GRID}[_Key]")
    first, last = top + 1, top + n_rows
    for formula, key, kind in (("=res_Name", "Resource", "list"), ("=leave_Months", "Month", "list")):
        dv = DataValidation(type=kind, formula1=formula, showErrorMessage=True, showInputMessage=True,
                            promptTitle=key, prompt=cols[0 if key == "Resource" else 1].note)
        ws.add_data_validation(dv)
        dv.add(col_range(letters[key], first, last))
    amount = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", showErrorMessage=True,
                            error="1 = a day off, 0.5 = half a day.", showInputMessage=True, promptTitle="Day off",
                            prompt="1 = all day, 0.5 = half a day. Weekends and holidays don't count.")
    ws.add_data_validation(amount)
    amount.add(f"{letters['1']}{first}:{letters['31']}{last}")
    return last + 1
