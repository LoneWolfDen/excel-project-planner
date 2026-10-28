"""Check sheet (Release 1.0, optional module "check"; not in Simple): everything to fix before the status report
goes out, on one page. A count per kind of problem with one verdict ("All clear" or how many to fix), then the
lists: tasks whose Issues column says something, over-allocated tasks and overdue Log items.

The lists are ranked tables (core.ranked_table), as on Status: each reads a key column on Tasks or the Log that is
blank for rows not on the list, so no FILTER (Microsoft 365 only) is needed and nothing is volatile. Delete the
sheet and nothing else changes: no other sheet reads it.
"""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font

from .core import (BOX, CALC_FILL, CALC_FONT, LABEL_FONT, NOTE_FONT, TITLE_FONT, Show, fill, font, has, header_cell,
                   name, ranked_table, v11, m365)

SHEET = "Check"
ROWS = 20                     # rows per list; drag a table's corner down for more

# label, count formula, where to fix it
COUNTS = [
    ("Tasks with something in Issues", "COUNT(tsk_IssueKey)", "Tasks > Issues: each says what is wrong"),
    ("Tasks sharing a UID", "DuplicateUIDs", "Give one of them Settings > NextFreeUID"),
    ("Named tasks without a UID", "MissingUIDs", "Type Settings > NextFreeUID"),
    ("Ignored predecessor links", 'COUNTIF(tsk_LinkCheck,"?*")', "Tasks > Link Check (DETAIL section)"),
    ("Over-allocated tasks", "COUNT(tsk_OverAllocKey)", "Resource Report: move work or add people"),
    ("Overdue Log items", "COUNT(log_OverdueKey)", "Log: update the item, its Due date or close it"),
    ("No status date typed", "IF(N(StatusDateInput)=0,1,0)", "Status: type the status date at the top"),
]
# template 1.1
V11_COUNTS = [
    ("Sandbox is on (what-if dates)", 'IF(Sandbox="Yes",1,0)', "Settings > Sandbox: set No before reporting"),
    ("External dependencies arriving too late", 'COUNTIF(log_DepGap,"<0")', "Log: Expected date, or move the task"),
]


def build_check(wb):
    ws = wb.create_sheet(SHEET)
    ws["B1"] = '=ProjectName&"  -  Check before you report"'
    ws["B1"].font = TITLE_FONT
    ws["B2"] = ("Everything to fix before the status report goes out. It updates as you edit; an empty list means "
                "nothing to do. Fix things on the sheet named on the right, not here.")
    ws["B2"].font = NOTE_FONT
    for col, width in {"A": 3, "B": 34, "C": 10, "D": 46, "E": 14, "F": 14, "G": 60}.items():
        ws.column_dimensions[col].width = width
    counts = [c for c in COUNTS + (V11_COUNTS if v11() else []) if has("log") or "log_" not in c[1]]
    if m365():                                         # no key columns: count the tasks and items directly
        from .m365 import COUNTS as DIRECT
        counts = [(label, DIRECT.get(f, f), where) for label, f, where in counts]
    top = 4
    for c, text in enumerate(("Check", "Count", "Where to fix it"), start=2):
        header_cell(ws.cell(top, c), text)
    for i, (label, formula, where) in enumerate(counts, start=1):
        ws.cell(top + i, 2, label).border = BOX
        cell = ws.cell(top + i, 3, f"={formula}")
        cell.font, cell.fill, cell.border, cell.number_format = CALC_FONT, CALC_FILL, BOX, "0"
        ws.cell(top + i, 4, where).font = NOTE_FONT
    first, last = top + 1, top + len(counts)
    name(wb, "chk_Counts", f"{SHEET}!$C${first}:$C${last}")
    verdict = ws.cell(top + len(counts) + 2, 2, '=IF(SUM(chk_Counts)=0,"All clear: ready to report",'
                                                 'SUM(chk_Counts)&" to fix before you report")')
    verdict.font = font(size=14, bold=True)
    name(wb, "chk_Verdict", f"{SHEET}!$B${verdict.row}")
    ws.conditional_formatting.add(f"C{first}:C{last}", FormulaRule(formula=[f"C{first}>0"], font=Font(
        color="9C0006", bold=True), fill=fill("FFC7CE")))
    ws.conditional_formatting.add(verdict.coordinate, FormulaRule(formula=[f"SUM(chk_Counts)>0"], font=Font(
        color="C00000", size=14, bold=True)))
    ws.conditional_formatting.add(verdict.coordinate, FormulaRule(formula=[f"SUM(chk_Counts)=0"], font=Font(
        color="548235", size=14, bold=True)))

    at = verdict.row + 3
    if m365():                                         # F-1108: FILTER lists instead of ranked tables
        from .m365 import check_lists, note
        note(ws, at - 2)
        check_lists(ws, at, ROWS)
    else:
        _ranked_lists(ws, at)
    for col, width in {"A": 3, "B": 34, "C": 10, "D": 46, "E": 14, "F": 14, "G": 60}.items():
        ws.column_dimensions[col].width = width     # the lists share the page's columns
    ws.freeze_panes = "A4"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def _ranked_lists(ws, at):
    """Release 1.0 / 1.1: ranked tables over the key columns (Excel 2019)."""
    task = Show("Task", "tsk_Name", width=46)
    ranked_table(ws, "CheckIssues", at, 1, "tsk_IssueKey", ROWS,
                 [Show("ID", "tsk_ID", "num", 34), Show("UID", "tsk_UID", "num", 10), task,
                  Show("Issues", "tsk_Issues", width=60)], label='"TASKS WITH ISSUES"')
    at += ROWS + 3
    ranked_table(ws, "CheckOverAllocated", at, 1, "tsk_OverAllocKey", ROWS,
                 [Show("ID", "tsk_ID", "num"), Show("UID", "tsk_UID", "num"), task,
                  Show("Start", "tsk_FcStart", "date", fmt="dd-mmm-yy"),
                  Show("Finish", "tsk_FcFinish", "date", fmt="dd-mmm-yy"),
                  Show("Resource Names", "tsk_ResourceNames")], label='"OVER-ALLOCATED TASKS"')
    if has("log"):
        at += ROWS + 3
        ranked_table(ws, "CheckOverdue", at, 1, "log_OverdueKey", ROWS,
                     [Show("Ref", "log_Ref", "num"), Show("Type", "log_Type"), Show("Title", "log_Title"),
                      Show("Due", "log_Due", "date", fmt="dd-mmm-yy"), Show("Owner", "log_Owner")],
                     label='"OVERDUE LOG ITEMS"')
