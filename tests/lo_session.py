"""Drive LibreOffice through UNO to edit a workbook the way a person would (insert or delete real rows).

The unit tests otherwise edit input cells with openpyxl, which can't show how Excel and LibreOffice adjust existing
formulas when rows move. Needs LibreOffice and python3-uno; tests skip without them.
"""
import itertools
import os
import subprocess
import time
from pathlib import Path

try:
    import uno
    from com.sun.star.beans import PropertyValue
    from com.sun.star.table import CellAddress, CellRangeAddress
    AVAILABLE = True
except ImportError:                      # pragma: no cover - depends on the machine
    AVAILABLE = False

# parallel workers (scripts/run_tests.py) each get their own range of ports and their own profile
_ports = itertools.count(int(os.environ.get("PLANNER_LO_PORT", "2101")))


def lo_profile():
    """-env:UserInstallation for a parallel worker's own LibreOffice profile (PLANNER_LO_PROFILE; see
    planner.build.lo_profile, which this mirrors so this module needs no planner import)."""
    folder = os.environ.get("PLANNER_LO_PROFILE")
    return [f"-env:UserInstallation={Path(folder).resolve().as_uri()}"] if folder else []


def _prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


class Calc:
    """Context manager: open a workbook in a headless LibreOffice, edit it, save a copy."""

    def __init__(self, path):
        self.path = Path(path).resolve()

    def __enter__(self):
        port = next(_ports)
        self.proc = subprocess.Popen(["soffice", *lo_profile(), "--headless", "--invisible", "--norestore", "--nologo",
                                      f"--accept=socket,host=localhost,port={port};urp;"],
                                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        ctx = uno.getComponentContext()
        resolver = ctx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", ctx)
        for _ in range(240):
            try:
                remote = resolver.resolve(f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(0.5)
        else:
            raise RuntimeError("LibreOffice did not start")
        self.desktop = remote.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", remote)
        self.doc = self.desktop.loadComponentFromURL(uno.systemPathToFileUrl(str(self.path)), "_blank", 0,
                                                     (_prop("Hidden", True),))
        return self

    def sheet(self, name):
        return self.doc.Sheets.getByName(name)

    def insert_row_like_excel(self, sheet_name, excel_row, ncols):
        """Insert a sheet row above excel_row and fill it with the formulas of the row above, as Excel does for a
        table's calculated columns. Input cells are cleared afterwards by the caller."""
        sh = self.sheet(sheet_name)
        index = [self.doc.Sheets.getByIndex(i).Name for i in range(self.doc.Sheets.Count)].index(sheet_name)
        sh.Rows.insertByIndex(excel_row - 1, 1)
        src = CellRangeAddress()
        src.Sheet, src.StartColumn, src.EndColumn = index, 0, ncols - 1
        src.StartRow = src.EndRow = excel_row - 2
        dst = CellAddress()
        dst.Sheet, dst.Column, dst.Row = index, 0, excel_row - 1
        sh.copyRange(dst, src)

    def delete_row(self, sheet_name, excel_row):
        self.sheet(sheet_name).Rows.removeByIndex(excel_row - 1, 1)

    def set(self, sheet_name, col, excel_row, value):
        cell = self.sheet(sheet_name).getCellByPosition(col, excel_row - 1)
        if value is None:
            cell.setString("")
        elif isinstance(value, str):
            cell.setString(value)
        else:
            cell.setValue(value)

    def save(self, out):
        self.doc.calculateAll()
        self.doc.storeToURL(uno.systemPathToFileUrl(str(Path(out).resolve())), (_prop("FilterName", "Calc Office Open XML"),))

    def __exit__(self, *exc):
        try:
            self.doc.close(True)
            self.desktop.terminate()
        except Exception:
            pass
        try:
            self.proc.wait(timeout=30)
        except subprocess.TimeoutExpired:
            self.proc.kill()
