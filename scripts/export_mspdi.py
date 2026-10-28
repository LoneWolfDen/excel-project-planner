"""Export a planner workbook to Microsoft Project XML (open the .xml in MS Project with File > Open).

    python scripts/export_mspdi.py Project_Planner.xlsx project.xml

The same conversion runs in the browser: docs/export.html (see planner/mspdi.py).
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from planner.mspdi import ExportError, convert_with_notes  # noqa: E402


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("workbook", help="planner .xlsx, saved after its last edit")
    p.add_argument("out", nargs="?", help="output .xml (default: the workbook name with .xml)")
    a = p.parse_args(argv)
    out = Path(a.out or Path(a.workbook).with_suffix(".xml"))
    try:
        xml, notes = convert_with_notes(Path(a.workbook).read_bytes())
    except ExportError as e:
        raise SystemExit(f"error: {e}")
    out.write_text(xml, encoding="utf-8")
    print("saved", out)
    for note in notes:                              # what MS Project won't see (B-1014)
        print("note:", note)


if __name__ == "__main__":
    main()
