"""Build the planner workbook.

    python scripts/build_planner.py Project_Planner.xlsx --recalc
    python scripts/build_planner.py Project_Planner_Blank.xlsx --blank --recalc
    python scripts/build_planner.py stress.xlsx --tasks 1000 --stress 33      # performance testing
"""
import argparse
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

import openpyxl

from . import sample
from .dashboard import build_dashboard
from .gantt import build_gantt
from .instructions import build_instructions
from .msp import build_msp_resources, build_msp_tasks
from .changes import build_changes
from .decisions import build_decisions
from .history import build_history
from .raid import build_raid
from .report_agile import build_agile_report
from .report_resources import build_resource_report
from .report_schedule import build_schedule_report
from .report_status import build_status_report
from .resources import build_resources
from .settings import build_holidays, build_lists, build_settings
from .sprints import build_sprints
from .tasks import build_tasks

ORDER = ["Instructions", "Dashboard", "Status Report", "Schedule Report", "Resource Report", "Agile Report", "RAID",
         "Decisions", "Changes", "History", "Settings", "Tasks", "Gantt", "Resources", "Sprints", "Holidays", "MSP_Tasks",
         "MSP_Resources", "Lists"]
TAB_COLOURS = {"Instructions": "1F4E78", "Dashboard": "548235", "Status Report": "548235",
               "Schedule Report": "548235", "Resource Report": "548235", "Agile Report": "548235", "RAID": "C55A11",
               "Decisions": "C55A11", "Changes": "C55A11", "History": "C55A11", "Settings": "BF8F00", "Tasks": "2F75B5",
               "Gantt": "2F75B5", "Resources": "7030A0", "Sprints": "7030A0", "Holidays": "BF8F00",
               "MSP_Tasks": "595959", "MSP_Resources": "595959", "Lists": "A6A6A6"}
N_RESOURCES, N_SPRINTS, N_GANTT_DAYS = 50, 26, 150


def make_workbook(blank=False, n_tasks=200, stress=0):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    rows = sample.tasks(blank, stress)
    if len(rows) > n_tasks:
        raise SystemExit(f"{len(rows)} sample tasks do not fit in {n_tasks} rows; raise --tasks")
    build_lists(wb)
    build_settings(wb, blank, N_SPRINTS)
    build_holidays(wb, blank)
    build_resources(wb, blank, N_RESOURCES)
    build_tasks(wb, rows, n_tasks)
    sprints_ws, sprint_letters = build_sprints(wb, blank, N_SPRINTS, 25)
    build_gantt(wb, n_tasks, N_GANTT_DAYS)
    build_msp_tasks(wb, n_tasks)
    build_msp_resources(wb, N_RESOURCES)
    build_dashboard(wb, sprints_ws, sprint_letters)
    build_raid(wb, blank)
    build_decisions(wb, blank)
    build_changes(wb, blank)
    build_status_report(wb)
    build_schedule_report(wb)
    build_resource_report(wb, 25)
    build_agile_report(wb, sprints_ws, sprint_letters)
    build_history(wb, blank)
    build_instructions(wb, n_tasks)
    wb._sheets = [wb[n] for n in ORDER]
    for ws in wb.worksheets:
        ws.sheet_properties.tabColor = TAB_COLOURS[ws.title]
    wb.active = ORDER.index("Tasks")
    wb.calculation.fullCalcOnLoad = True
    return wb


def _tables(z):
    """{displayName: (path, xml)} for every table part in an .xlsx zip."""
    out = {}
    for n in z.namelist():
        if n.startswith("xl/tables/") and n.endswith(".xml"):
            xml = z.read(n).decode("utf-8")
            out[re.search(r'displayName="([^"]+)"', xml).group(1)] = (n, xml)
    return out


def restore_tables(original, recalculated):
    """LibreOffice drops each table's calculated-column formulas when it saves. Excel needs them to fill formulas
    into newly inserted rows, so put the original table definitions back (keeping LibreOffice's part ids)."""
    with zipfile.ZipFile(original) as z:
        source = _tables(z)
    tmp = Path(str(recalculated) + ".tmp")
    with zipfile.ZipFile(recalculated) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        target = {path: name for name, (path, _) in _tables(zin).items()}
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in target and target[item.filename] in source:
                lo_id = re.search(r'<table[^>]*\sid="(\d+)"', data.decode("utf-8")).group(1)
                xml = source[target[item.filename]][1]
                data = re.sub(r'(<table[^>]*\s)id="\d+"', rf'\g<1>id="{lo_id}"', xml, count=1).encode("utf-8")
            zout.writestr(item, data)
    tmp.replace(recalculated)


def save(wb, out, recalc):
    if not recalc:
        wb.save(out)
        return
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise SystemExit("--recalc needs LibreOffice (soffice) on PATH")
    with tempfile.TemporaryDirectory() as tmp:
        raw = Path(tmp) / Path(out).name
        wb.save(raw)
        subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(Path(tmp) / "out"),
                        str(raw)], check=True, capture_output=True, timeout=1800)
        done = Path(tmp) / "out" / raw.name
        restore_tables(raw, done)
        shutil.copyfile(done, out)


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("out", help="output .xlsx path")
    p.add_argument("--blank", action="store_true", help="empty template: no sample tasks, resources, holidays")
    p.add_argument("--recalc", action="store_true",
                   help="recalculate with LibreOffice so the saved file carries values (for previews)")
    p.add_argument("--tasks", type=int, default=200, help="rows in the Tasks table (and Gantt / MSP_Tasks)")
    p.add_argument("--stress", type=int, default=0, help="repeat the sample project N times (performance tests)")
    a = p.parse_args(argv)
    save(make_workbook(a.blank, a.tasks, a.stress), a.out, a.recalc)
    print("saved", a.out)


if __name__ == "__main__":
    main()
