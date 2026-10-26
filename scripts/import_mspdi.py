"""Import a Microsoft Project XML file (MS Project: File > Save As > XML) into a new planner workbook.

    python scripts/import_mspdi.py project.xml Project_Planner.xlsx [--recalc]

The same conversion runs in the browser: docs/export.html (see planner/mspdi_import.py). Anything the planner can't
hold is listed (and noted on the task).
"""
import argparse
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import openpyxl  # noqa: E402

from planner.build import save  # noqa: E402
from planner.mspdi_import import MSPImportError, convert  # noqa: E402


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("xml", help="MS Project XML (MSPDI) file")
    p.add_argument("out", nargs="?", help="output .xlsx (default: the XML name with .xlsx)")
    p.add_argument("--recalc", action="store_true",
                   help="recalculate with LibreOffice so the file carries values (Excel recalculates on open anyway)")
    a = p.parse_args(argv)
    out = Path(a.out or Path(a.xml).with_suffix(".xlsx"))
    try:
        xlsx, warnings = convert(Path(a.xml).read_bytes())
    except MSPImportError as e:
        raise SystemExit(f"error: {e}")
    if a.recalc:
        with tempfile.TemporaryDirectory() as tmp:
            raw = Path(tmp) / out.name
            raw.write_bytes(xlsx)
            save(openpyxl.load_workbook(raw), out, True)
    else:
        out.write_bytes(xlsx)
    for w in warnings:
        print("note:", w)
    print("saved", out)


if __name__ == "__main__":
    main()
