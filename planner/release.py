"""Release 1.0 finishing touches, applied once every sheet is built (build.make_workbook, --release only).

  * TemplateVersion: the release users see, a defined name ("1.0") shown on Start and in every printed footer.
    PlannerVersion (2.0) stays the code's version, which planner/upgrade.py reads.
  * Printing: Status, Roadmap, Resource Report and Check print on one page wide (Status and Roadmap on one page)
    with the sheet and file name in the header and the template version and page number in the footer. Excel's
    headers and footers can't read a cell, so the project name and status date are printed from each sheet's title
    rows, which sit inside the print area.
  * Protection-ready input sheets: on Tasks, Resources, Settings and History the blue input cells are unlocked, so
    Review > Protect Sheet (no password) locks only the formulas. They are left unprotected because Excel can't add
    table rows on a protected sheet; build.RELEASE_VIEWS lists the sheets that ship protected.
"""
import re

from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.styles import Protection
from openpyxl.utils.cell import range_boundaries

from .core import INPUT_FONT, NOTE_FONT, name, v11
from .settings import START

PRINTED = ("Status", "Roadmap", "Resource Report", "Check")
OPEN_INPUTS = ("Tasks", "Resources", "Settings", "History")


def _unlock_inputs(ws):
    blue = INPUT_FONT.color.rgb
    for row in ws.iter_rows():
        for cell in row:
            if cell.font is not None and cell.font.color is not None and cell.font.color.rgb == blue:
                cell.protection = Protection(locked=False)


def _print_setup(ws, version):
    ws.oddHeader.left.text = "&A"
    ws.oddHeader.right.text = "&F"
    ws.oddFooter.left.text = f"Project Planner template {version}"
    ws.oddFooter.right.text = "Page &P of &N"
    for part in (ws.oddHeader.left, ws.oddHeader.right, ws.oddFooter.left, ws.oddFooter.right):
        part.size, part.font = 8, "Arial"
    ws.print_options.horizontalCentered = True
    ws.page_margins.top = ws.page_margins.bottom = 0.6
    ws.page_margins.header = ws.page_margins.footer = 0.3


# a cell reference whose column isn't fixed ($): a rule using one means something different on each range
RELATIVE_COLUMN = re.compile(r"(?<![$A-Za-z0-9_\[\]])[A-Z]{1,3}\$?\d+\b")


def merge_rules(ws):
    """F-1109: one conditional-formatting rule where several are the same rule on different column ranges.

    Two rules merge only when their type, formula, format and operator are the same, their formula fixes every
    column ($X6), and their ranges start on the same row: Excel then reads the formula the same way on every range,
    whichever range it takes as the anchor. Rules keep the order (priority) of their first range."""
    merged, order = {}, []
    for cf in ws.conditional_formatting:
        for rule in cf.rules:
            formula = tuple(rule.formula or ())
            tops = {range_boundaries(r)[1] for r in str(cf.sqref).split()}
            fixed = not any(RELATIVE_COLUMN.search(f) for f in formula)
            key = (rule.type, formula, repr(rule.dxf), rule.operator, rule.stopIfTrue, min(tops)) if fixed and \
                len(tops) == 1 else (id(rule), str(cf.sqref))
            if key not in merged:
                merged[key] = (rule, [])
                order.append(key)
            merged[key][1].append(str(cf.sqref))
    rules = ConditionalFormattingList()
    for key in order:
        rule, ranges = merged[key]
        rules.add(" ".join(dict.fromkeys(" ".join(ranges).split())), rule)
    ws.conditional_formatting = rules


def finish_release(wb, version):
    name(wb, "TemplateVersion", f'"{version}"')
    if v11():
        for ws in wb.worksheets:
            merge_rules(ws)
    for sheet in PRINTED:
        if sheet in wb.sheetnames:
            _print_setup(wb[sheet], version)
    for sheet in ("Roadmap", "Check"):
        if sheet in wb.sheetnames:
            ws = wb[sheet]
            ws.print_area = f"A1:{ws.cell(ws.max_row, ws.max_column).coordinate}"
    if "Resource Report" in wb.sheetnames:            # one page: the title and the per-resource summary
        ws = wb["Resource Report"]
        ws["A1"] = '=ProjectName&"  -  Resource Report   (status date "&TEXT(StatusDate,"dd-mmm-yyyy")&")"'
        summary = ws.tables["ResourceSummary"].ref
        ws.print_area = f"A1:{summary.split(':')[1]}"
        ws.page_setup.fitToHeight = 1
    for sheet in OPEN_INPUTS:
        if sheet in wb.sheetnames:
            _unlock_inputs(wb[sheet])
    start = wb[START]
    r = start.max_row + 2
    start.cell(r, 2, '="Project Planner template "&TemplateVersion&"  -  Excel 2019 or Microsoft 365, desktop, '
                     'web and iPhone. No macros."').font = NOTE_FONT
    start.oddFooter.left.text = f"Project Planner template {version}"
