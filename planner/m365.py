"""The Microsoft 365 edition (F-1108, `--release m365`, release-1.1-m365/): template 1.1 rebuilt with the functions
Excel 2019 lacks, in its own folder so 1.1 is never at risk.

  * Named LAMBDA functions in Name Manager, so a formula reads like what it does:
      L_PARSE_DURATION(x)   working days from 3, "3d" or "2w" ("" when unreadable): Tasks' _Days and Sandbox
      L_WORKING_DAYS(a, b)  signed working days from a to b (WorkWeek, Holidays): Log > Gap, the finish variance
      L_IS_WORKING_DAY(d)   TRUE on a working day: Tasks > Issues flags actual dates on a weekend or holiday
    They are written with LET inside, and stored as Excel stores them (_xlfn.LAMBDA, _xlfn.LET, _xlpm. parameters).
  * The Check sheet's lists are FILTER results instead of ranked tables, so the key columns those needed
    (_IssueKey, _OverAllocKey, _OverdueKey) are gone. Each list is a fixed-size array formula (openpyxl can't write
    Excel's spilled-array metadata); the rows FILTER doesn't fill stay blank. Make a list longer with Formulas >
    (select the block, edit, Ctrl+Shift+Enter) or delete the block and type the formula once to let it spill.

The scheduling formulas themselves (the forward and backward passes) are the same as in 1.1, so the plan schedules
identically; only the places above use the new functions. LibreOffice 24.2, which runs the tests, can't calculate
LAMBDA or FILTER, so this edition is checked by structure in tests/test_m365.py and by hand in Excel
(docs/TESTING_IN_EXCEL.md, section 8). The workbook is saved without cached values: Excel calculates on open.
"""
from openpyxl.styles import Alignment
from openpyxl.worksheet.formula import ArrayFormula

from .core import BOX, CALC_FILL, CALC_FONT, LABEL_FONT, font, has, header_cell, name

P = "_xlpm."                              # how a LAMBDA / LET parameter is stored in the file
LAMBDAS = {
    "L_PARSE_DURATION": (
        f"_xlfn.LAMBDA({P}x,_xlfn.LET({P}t,LOWER(TRIM({P}x)),"
        f'IF({P}x="","",IF(ISNUMBER({P}x),{P}x,IFERROR(VALUE({P}t),'
        f'IFERROR(VALUE(LEFT({P}t,LEN({P}t)-1))*_xlfn.SWITCH(RIGHT({P}t,1),"w",5,"d",1),""))))))'),
    "L_WORKING_DAYS": (
        f"_xlfn.LAMBDA({P}a,{P}b,IF({P}b>={P}a,NETWORKDAYS.INTL({P}a,{P}b,WorkWeek,Holidays)-1,"
        f"1-NETWORKDAYS.INTL({P}b,{P}a,WorkWeek,Holidays)))"),
    "L_IS_WORKING_DAY": f"_xlfn.LAMBDA({P}d,NETWORKDAYS.INTL({P}d,{P}d,WorkWeek,Holidays)=1)",
}


def define(wb):
    for n, body in LAMBDAS.items():
        name(wb, n, body)


def days(raw):
    return f"L_PARSE_DURATION({raw})"


def working_days(a, b):
    return f"L_WORKING_DAYS({a},{b})"


def is_working_day(d):
    return f"L_IS_WORKING_DAY({d})"


def _filter(columns, condition):
    """FILTER of several named columns side by side (CHOOSE with an array of column numbers), blank when empty."""
    pick = "CHOOSE({" + ",".join(str(k) for k in range(1, len(columns) + 1)) + "}," + ",".join(columns) + ")"
    return f'=IFERROR(_xlfn._xlws.FILTER({pick},{condition}),"")'


def filter_list(ws, top, label, headers, columns, condition, rows, formats=None, count=None):
    """A heading, a header row (columns B onwards) and a `rows`-row array formula filling it. Returns the next row."""
    ws.cell(top - 1, 2, f'={label}&": "&{count}' if count else label).font = LABEL_FONT
    for k, h in enumerate(headers):
        header_cell(ws.cell(top, 2 + k), h)
    first, last = top + 1, top + rows
    end = ws.cell(last, 1 + len(headers)).column_letter
    ws[f"B{first}"] = ArrayFormula(f"B{first}:{end}{last}", _filter(columns, condition))
    for r in range(first, last + 1):
        for k in range(len(headers)):
            c = ws.cell(r, 2 + k)
            c.font, c.fill, c.border = CALC_FONT, CALC_FILL, BOX
            c.alignment = Alignment(vertical="top", wrap_text=k == len(headers) - 1)
            if formats and formats[k]:
                c.number_format = formats[k]
    return last + 3


def check_lists(ws, at, rows):
    """The Check sheet's three lists, as FILTERs."""
    at = filter_list(ws, at, '"TASKS WITH ISSUES"', ["ID", "UID", "Task", "Issues"],
                     ["tsk_ID", "tsk_UID", "tsk_Name", "tsk_Issues"], '(tsk_Issues<>"")*(tsk_Name<>"")', rows,
                     count='COUNTIFS(tsk_Issues,"?*",tsk_Name,"?*")')
    at = filter_list(ws, at, '"OVER-ALLOCATED TASKS"', ["ID", "UID", "Task", "Start", "Finish", "Resource Names"],
                     ["tsk_ID", "tsk_UID", "tsk_Name", "tsk_FcStart", "tsk_FcFinish", "tsk_ResourceNames"],
                     'tsk_OverAllocated="Yes"', rows, [None, None, None, "dd-mmm-yy", "dd-mmm-yy", None],
                     count='COUNTIF(tsk_OverAllocated,"Yes")')
    if has("log"):
        at = filter_list(ws, at, '"OVERDUE LOG ITEMS"', ["Ref", "Type", "Title", "Due", "Owner"],
                         ["log_Ref", "log_Type", "log_Title", "log_Due", "log_Owner"], 'log_Overdue="Overdue"', rows,
                         [None, None, None, "dd-mmm-yy", None], count='COUNTIF(log_Overdue,"Overdue")')
    return at


# the Check sheet's counts that read the key columns in the other editions
COUNTS = {"COUNT(tsk_IssueKey)": 'COUNTIFS(tsk_Issues,"?*",tsk_Name,"?*")',
          "COUNT(tsk_OverAllocKey)": 'COUNTIF(tsk_OverAllocated,"Yes")',
          "COUNT(log_OverdueKey)": 'COUNTIF(log_Overdue,"Overdue")'}


def note(ws, row):
    ws.cell(row, 2, "Microsoft 365 edition: these lists are FILTER formulas (they show the first rows that fit). "
                    "Needs Excel for Microsoft 365 or Excel 2021+.").font = font(italic=True, color="595959")
