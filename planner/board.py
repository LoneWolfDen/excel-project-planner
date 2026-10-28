"""Board sheet (template 1.2, B-404; optional module "board", not in Simple): the detail tasks as cards in five
columns, Not Started, Late Start, In Progress, Behind and Complete, filtered by sprint and by person.

Each card cell finds its task without any per-column helper on Tasks: Tasks' _BoardKey is status number x 100000 +
the task's row (blank when the filters leave it out), so column s holds the keys between s x 100000 and the next
status, and card k of column s is the (COUNTIF of the keys below s x 100000 + k)-th smallest key. One hidden
position cell per card does that MATCH once; the card shows the task's name and people. Cells are found by
ROW() offsets from the named anchor brd_Origin (each status's position cells are named brd_Pos1 ... brd_Pos5), so
nothing uses an A1 address.
"""
from openpyxl.styles import Alignment
from openpyxl.worksheet.datavalidation import DataValidation

from .core import BOX, CALC_FONT, INPUT_FONT, LABEL_FONT, NOTE_FONT, TITLE_FONT, fill, font, header_cell, name
from .tasks import BOARD_STATUSES, STATUS_COLOURS

SHEET = "Board"
CARDS = 25                       # cards per column
TOP = 7                          # the status headers' row


def build_board(wb):
    ws = wb.create_sheet(SHEET)
    ws["B1"] = '=ProjectName&"  -  Board"'
    ws["B1"].font = TITLE_FONT
    ws["B2"] = ("Every detail task as a card, by status. Pick a sprint or a person to narrow it down; blank = all. "
                "Change a task on Tasks (its Stage and actual dates) and its card moves.")
    ws["B2"].font = NOTE_FONT
    for r, (label, n, source, prompt) in enumerate((
            ("Sprint", "BoardSprint", "=spr_Name", "Show the tasks of one sprint (blank = every sprint)."),
            ("Person", "BoardResource", "=res_Name", "Show one person's tasks (blank = everyone).")), start=4):
        ws.cell(r, 2, label).font = LABEL_FONT
        cell = ws.cell(r, 3)
        cell.font, cell.border = INPUT_FONT, BOX
        name(wb, n, f"{SHEET}!$C${r}")
        dv = DataValidation(type="list", formula1=source, showErrorMessage=False, showInputMessage=True,
                            promptTitle=label, prompt=prompt)
        ws.add_data_validation(dv)
        dv.add(cell.coordinate)
    if "spr_Name" not in wb.defined_names:             # no Agile sheet: no sprints to pick from
        ws.cell(4, 4, "(no sprint calendar in this workbook)").font = NOTE_FONT
    origin = 2                                         # brd_Origin: the column left of the first card column
    name(wb, "brd_Origin", f"{SHEET}!${ws.cell(TOP, origin).column_letter}${TOP}")
    ws.column_dimensions[ws.cell(1, origin).column_letter].width = 2
    k = "(ROW()-ROW(brd_Origin))"                     # card number in its column
    first, last = TOP + 1, TOP + CARDS
    for s, label in enumerate(BOARD_STATUSES, start=1):
        pos_col, card_col = origin + 2 * s - 1, origin + 2 * s
        pos_letter = ws.cell(1, pos_col).column_letter
        ws.column_dimensions[pos_letter].hidden = True
        ws.column_dimensions[ws.cell(1, card_col).column_letter].width = 30
        name(wb, f"brd_Pos{s}", f"{SHEET}!${pos_letter}${first}:${pos_letter}${last}")
        lo, hi = s * 100000, (s + 1) * 100000         # this status's keys
        head = ws.cell(TOP, card_col)
        header_cell(head, None)
        head.value = f'="{label}  ("&COUNTIFS(tsk_BoardKey,">={lo}",tsk_BoardKey,"<{hi}")&")"'
        head.fill = fill(STATUS_COLOURS[label])
        nth = f'SMALL(tsk_BoardKey,COUNTIF(tsk_BoardKey,"<{lo}")+{k})'
        for r in range(first, last + 1):
            ws.cell(r, pos_col, f'=IFERROR(IF({nth}<{hi},MATCH({nth},tsk_BoardKey,0),""),"")').font = CALC_FONT
            pos = f"INDEX(brd_Pos{s},{k})"
            card = ws.cell(r, card_col, f'=IF({pos}="","",INDEX(tsk_Name,{pos})&IF(INDEX(tsk_ResourceNames,{pos})'
                                        f'="","",CHAR(10)&INDEX(tsk_ResourceNames,{pos})))')
            card.font, card.border = CALC_FONT, BOX
            card.alignment = Alignment(wrap_text=True, vertical="top")
    for r in range(first, last + 1):
        ws.row_dimensions[r].height = 30
    ws.freeze_panes = ws.cell(first, 1)
    ws.sheet_view.zoomScale = 100
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws

