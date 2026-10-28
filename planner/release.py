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
from openpyxl.styles import Protection

from .core import INPUT_FONT, NOTE_FONT, name
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


def finish_release(wb, version):
    name(wb, "TemplateVersion", f'"{version}"')
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
