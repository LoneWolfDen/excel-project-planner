"""Shared building blocks: styles, defined names and Excel Tables.

Every formula the planner writes refers to cells through Excel Tables (structured references such as
``Tasks[[#This Row],[Duration]]``, shown by Excel as ``[@Duration]``) or defined names (``ProjectStart``,
``tsk_Duration``). No formula contains an A1 address, so inserting or moving rows and columns never breaks it.
Conditional formatting, data validation and charts are the only places that use A1 ranges, because Excel does
not accept structured references there; Excel shifts those ranges itself when rows or columns are inserted.
"""
from dataclasses import dataclass, field

from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.table import Table, TableColumn, TableFormula, TableStyleInfo

FONT = "Arial"
DATE = "ddd dd-mmm-yy"
SHORT_DATE = "dd-mmm-yy"


def font(**k):
    k.setdefault("name", FONT)
    k.setdefault("size", 10)
    return Font(**k)


def fill(colour):
    return PatternFill("solid", start_color=colour, end_color=colour)


INPUT_FONT = font(color="0000FF")
CALC_FONT = font()
HELPER_FONT = font(color="404040")
CALC_FILL = fill("F2F2F2")
HEADER_FILLS = {"in": fill("1F4E78"), "calc": fill("595959"), "help": fill("A6A6A6")}
HEADER_FONT = font(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="D9D9D9")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
CENTER = Alignment(horizontal="center")
TITLE_FONT = font(size=16, bold=True, color="1F4E78")
NOTE_FONT = font(italic=True, color="595959")
LABEL_FONT = font(bold=True, color="1F4E78")


def name(wb, defined_name, ref):
    """Workbook-level defined name, e.g. name(wb, "ProjectStart", "Settings!$C$7")."""
    wb.defined_names[defined_name] = DefinedName(defined_name, attr_text=ref)


def title(ws, text, note=None):
    ws["A1"] = text
    ws["A1"].font = TITLE_FONT
    if note:
        ws["A2"] = note
        ws["A2"].font = NOTE_FONT


def header_cell(cell, text, kind="calc"):
    cell.value = text
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILLS[kind]
    cell.alignment = WRAP_CENTER
    cell.border = BOX


@dataclass
class Col:
    """One column of an Excel Table.

    name     header text, which is also the structured-reference name ([@name])
    kind     "in" (user input, blue), "calc" (formula, grey) or "help" (hidden helper formula)
    formula  formula for every data row (same text in every row: a calculated column), or None for inputs
    """
    name: str
    kind: str = "calc"
    width: float = 10
    fmt: str = None
    formula: str = None
    center: bool = False
    note: str = None
    key: str = None
    values: list = field(default_factory=list)   # optional constants for input columns, row by row


class Tab:
    """Structured-reference helpers for one Excel Table."""

    def __init__(self, table):
        self.table = table

    def row(self, col):       # the value in this row: shown by Excel as [@col]
        return f"{self.table}[[#This Row],[{col}]]"

    def col(self, col):       # the whole data column
        return f"{self.table}[{col}]"

    def head(self, col):      # the header cell, used as an anchor for OFFSET and ROW
        return f"{self.table}[[#Headers],[{col}]]"

    def upto(self, col):      # header cell down to this row (an expanding range that never includes later rows)
        return f"{self.head(col)}:{self.row(col)}"


def add_table(ws, table_name, header_row, first_col, cols, n_rows, style="TableStyleLight1", autofilter=True):
    """Write a styled Excel Table with calculated columns. Returns {column name: column letter}."""
    letters = {}
    last_row = header_row + n_rows
    for i, c in enumerate(cols):
        col = first_col + i
        letter = get_column_letter(col)
        letters[c.name] = letter
        if c.key:
            letters[c.key] = letter
        header_cell(ws.cell(header_row, col), c.name, c.kind)
        if c.note:
            ws.cell(header_row, col).comment = Comment(c.note, "Planner")
        ws.column_dimensions[letter].width = c.width
        for r in range(header_row + 1, last_row + 1):
            cell = ws.cell(r, col)
            cell.border = BOX
            if c.kind == "in":
                cell.font = INPUT_FONT
                k = r - header_row - 1
                if k < len(c.values) and c.values[k] is not None:
                    cell.value = c.values[k]
            else:
                cell.font = HELPER_FONT if c.kind == "help" else CALC_FONT
                cell.fill = CALC_FILL
                k = r - header_row - 1
                if c.formula:
                    cell.value = f"={c.formula}"
                elif k < len(c.values):        # per-row content (labels or formulas), not a calculated column
                    cell.value = c.values[k]
            if c.fmt:
                cell.number_format = c.fmt
            if c.center:
                cell.alignment = CENTER
    ref = f"{get_column_letter(first_col)}{header_row}:{get_column_letter(first_col + len(cols) - 1)}{last_row}"
    table = Table(displayName=table_name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=False)
    if autofilter:
        table.autoFilter = AutoFilter(ref=ref)
    table.tableColumns = [
        TableColumn(id=i + 1, name=c.name,
                    calculatedColumnFormula=TableFormula(attr_text=c.formula) if c.formula else None)
        for i, c in enumerate(cols)
    ]
    ws.add_table(table)
    return letters


def col_range(letter, first_row, last_row):
    return f"{letter}{first_row}:{letter}{last_row}"
