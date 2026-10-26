"""Leave sheet: each person's leave, one column per month (Excel Table "LeaveTable"), and the formula that reads it.

A cell under "Nov-26" holding "16, 17(0.5)" means a full day off on 16-Nov-26 and half a day on 17-Nov-26
(semicolons work too; spaces are ignored). Leave reduces capacity only: utilisation, over-allocation and sprint
capacity count it, task dates don't move (the plan stays the same in MS Project).

Reading the text needs no helper cells:
  * the month's column is found by its header text, so columns can be added (type "Jan-29") or moved. The label is
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

from .core import Col, add_table, col_range, name, title

TABLE = "LeaveTable"
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
    return f"IFERROR(MATCH({label},{TABLE}[#Headers],0),0)"


def leave_row(resource):
    """Position of a person's row in LeaveTable; 0 when they have none."""
    return f"IFERROR(MATCH({resource},{TABLE}[Resource],0),0)"


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
    terms = [f'IF(OR({row}=0,{col}=0),0,{_in_month(start, end, first, _text(f"INDEX({TABLE}[#Data],{row},{col})"))})'
             for first, col in columns]
    return "(" + "+".join(terms) + ")"


def team_leave_hours(grid, column):
    """Capacity hours the team's leave removes in one period: the sum over a per-person leave grid's column (people
    in ResourceTable order) of days x Max Units x HoursPerDay, for Work resources. Reading the grid, which already
    holds each person's days, is far cheaper than parsing everyone's leave text again."""
    people = f"ROWS({grid})"
    units = f"INDEX(res_MaxUnits,1):INDEX(res_MaxUnits,{people})"
    work = f'(INDEX(res_Type,1):INDEX(res_Type,{people})="Work")'
    return f"(HoursPerDay*SUMPRODUCT({work}*{units}*INDEX({grid},0,{column})))"


def build_leave(wb, blank, n_rows=50, months=None, entries=None):
    """entries: {person: {month header: days text}} replacing the sample (an import); months: the month headers."""
    ws = wb.create_sheet("Leave")
    title(ws, "Leave",
          'One row per person, one column per month. Type the days off, e.g. "16, 17(0.5)" under Nov-26 = all of '
          "16 November and half of the 17th. Leave lowers capacity (Resource Report, Sprints) but doesn't move tasks. "
          'Add a month by typing its name (e.g. "Jan-29") in the header cell after the last month.')
    ws.column_dimensions["A"].width = 3
    labels = months or month_labels()
    data = entries if entries is not None else {} if blank else SAMPLE
    people = list(data)
    cols = [Col("Resource", "in", 20, values=people, note="Name as on the Resources sheet.")]
    for label in labels:
        cols.append(Col(label, "in", 11, values=[data[p].get(label) for p in people],
                        note=f'Days off in {label}: "16, 17(0.5)" = 16th all day, 17th half a day.'))
    letters = add_table(ws, TABLE, 4, 2, cols, n_rows)
    name(wb, "leave_Days", "{" + ",".join(str(d) for d in range(1, 32)) + "}")
    rule = DataValidation(type="list", formula1="=res_Name")
    ws.add_data_validation(rule)
    rule.add(col_range(letters["Resource"], 5, 4 + n_rows))
    ws.freeze_panes = "C5"
    return ws
