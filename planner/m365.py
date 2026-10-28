"""The Microsoft 365 edition (F-1108, `--release m365`, release-1.1-m365/): template 1.1 rebuilt with the functions
Excel 2019 lacks, in its own folder so 1.1 is never at risk.

  * Named LAMBDA functions in Name Manager, so a formula reads like what it does:
      L_PARSE_DURATION(x)   working days from 3, "3d" or "2w" ("" when unreadable): Tasks' _Days and Sandbox
      L_WORKING_DAYS(a, b)  signed working days from a to b (WorkWeek, Holidays): Log > Gap, the finish variance
      L_IS_WORKING_DAY(d)   TRUE on a working day: Tasks > Issues flags actual dates on a weekend or holiday
    They are written with LET inside, and stored as Excel stores them (_xlfn.LAMBDA, _xlfn.LET, _xlpm. parameters).
  * The Check sheet's lists are FILTER results instead of ranked tables, so the key columns those needed
    (_IssueKey, _OverAllocKey, _OverdueKey) are gone. Each list is one formula that spills (F-1110): openpyxl can't
    write Excel's dynamic-array metadata, so save() adds it to the file (xl/metadata.xml and cm="1" on the cell),
    which is what Excel itself writes for a formula typed in Microsoft 365. The rows below are formatted for the
    first `rows` results; a longer list spills further.

The scheduling formulas themselves (the forward and backward passes) are the same as in 1.1, so the plan schedules
identically; only the places above use the new functions. LibreOffice 24.2, which runs the tests, can't calculate
LAMBDA or FILTER, so this edition is checked by structure in tests/test_m365.py and by hand in Excel
(docs/TESTING_IN_EXCEL.md, section 8). The workbook is saved without cached values: Excel calculates on open.
"""
import re
import zipfile
from pathlib import Path

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
    ws[f"B{first}"] = ArrayFormula(f"B{first}", _filter(columns, condition))
    spill(ws, f"B{first}")
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
    ws.cell(row, 2, "Microsoft 365 edition: each list is one FILTER formula that spills as long as it needs. "
                    "Needs Excel for Microsoft 365 or Excel 2021+.").font = font(italic=True, color="595959")


# F-1110: Excel's dynamic-array metadata. A cell with cm="1" points at this cell-metadata record, which marks its
# array formula as dynamic (it spills) rather than a legacy Ctrl+Shift+Enter block.
METADATA = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
            '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" '
            'pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" '
            'coerce="1" cellMeta="1"/></metadataTypes>'
            '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
            '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>'
            '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>')
META_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"
META_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"


def spill(ws, cell):
    """Mark a single-cell array formula as one that spills, and make the workbook add the metadata when saved."""
    wb = ws.parent
    if not hasattr(wb, "spills"):
        wb.spills = []
        plain = wb.save

        def save(filename):
            plain(filename)
            if isinstance(filename, (str, Path)):
                add_spill_metadata(filename, wb.spills)
        wb.save = save
    wb.spills.append((ws.title, cell))


def add_spill_metadata(path, spills):
    """Rewrite the saved file: xl/metadata.xml, its content type and relationship, and cm="1" on each spill cell."""
    path = Path(path)
    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(path) as zin:
        rels = zin.read("xl/_rels/workbook.xml.rels").decode()
        book = zin.read("xl/workbook.xml").decode()
        target = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
        target.update({i: t for t, i in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
        sheets = {name: "xl/" + target[rid].lstrip("/").removeprefix("xl/")
                  for name, rid in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', book)}
        cells = {}
        for title, cell in spills:
            cells.setdefault(sheets[title], []).append(cell)
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename in cells:
                    xml = data.decode()
                    for cell in cells[item.filename]:
                        xml = re.sub(rf'<c r="{cell}"(?![^>]*\scm=)', f'<c r="{cell}" cm="1"', xml, count=1)
                    data = xml.encode()
                elif item.filename == "[Content_Types].xml" and "/xl/metadata.xml" not in data.decode():
                    data = data.decode().replace(
                        "</Types>", f'<Override PartName="/xl/metadata.xml" ContentType="{META_TYPE}"/></Types>').encode()
                elif item.filename == "xl/_rels/workbook.xml.rels" and META_REL not in data.decode():
                    data = data.decode().replace(
                        "</Relationships>",
                        f'<Relationship Id="rIdSpill" Type="{META_REL}" Target="metadata.xml"/></Relationships>').encode()
                zout.writestr(item, data)
            if "xl/metadata.xml" not in zin.namelist():
                zout.writestr("xl/metadata.xml", METADATA)
    tmp.replace(path)
