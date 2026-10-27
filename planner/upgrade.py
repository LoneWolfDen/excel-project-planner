"""Move a plan into the latest template (B-901): read what the user typed into an older planner workbook and build
the current workbook around it.

Only typed values are read (formula cells are skipped), and always by header name, never by position, so a column
that moved or a sheet that was reordered still lines up. The plan's core (tasks, people, assignments, holidays,
leave, settings) goes through planner.build.ProjectData, as an MS Project import does, so the new workbook's rows,
UIDs and sizes are worked out by the builder. Every other input table (the logs, status history, sprint goals,
PM RAG and notes, resource notes) is then copied onto the new workbook column by column, matched by table and
header. Dropdown lists keep the new template's values and add any the user had added.

Inputs that a version renamed or moved are mapped in RENAMED and in the per-version steps below; workbooks carry
their version in the PlannerVersion name (absent before v1.7). The same code runs on the command line
(scripts/upgrade_planner.py) and in the browser page docs/export.html.
"""
import io
import re

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries

from .core import BOX, HEADER_FILLS, INPUT_FONT

# Tasks headers an older version used, and the column they are now
RENAMED = {"Tasks": {"Start No Earlier Than": "Constraint Date"}}
# tables the builder fills from ProjectData; everything else with inputs is copied by header
BUILT = {"Tasks", "HolidayTable", "LeaveTable", "AssignmentTable"}
INPUT_FILL = HEADER_FILLS["in"].start_color.rgb[-6:]
MONTH = re.compile(r"^[A-Z][a-z]{2}-\d\d$")          # a leave month header: Nov-26


class UpgradeError(Exception):
    """A file that isn't a planner workbook (shown as-is by the command line and the browser page)."""


def _tables(wb):
    """{table name: (sheet, [header], [[typed value or None per column] per row])}, formulas left out."""
    out = {}
    for ws in wb.worksheets:
        for name, table in ws.tables.items():
            c1, r1, c2, r2 = range_boundaries(table.ref if hasattr(table, "ref") else table)
            head = [ws.cell(r1, c).value for c in range(c1, c2 + 1)]
            rows = []
            for r in range(r1 + 1, r2 + 1):
                row = []
                for c in range(c1, c2 + 1):
                    v = ws.cell(r, c).value
                    row.append(None if isinstance(v, str) and v.startswith("=") else v)
                rows.append(row)
            out[name] = (ws, head, rows)
    return out


def _records(tables, name, renamed=None):
    """A table's rows as {header: typed value}, trailing empty rows dropped."""
    if name not in tables:
        return []
    _, head, rows = tables[name]
    head = [(renamed or {}).get(h, h) for h in head]
    records = [{h: v for h, v in zip(head, row) if h is not None} for row in rows]
    while records and all(v in (None, "") for v in records[-1].values()):
        records.pop()
    return records


def _named_value(wb, name):
    if name not in wb.defined_names:
        return None
    ref = wb.defined_names[name].attr_text
    if "!" not in ref:                                   # a constant name, e.g. PlannerVersion = "1.7"
        return ref.strip('"') if ref.startswith('"') else None
    sheet, cell = ref.split("!")
    sheet = sheet.strip("'")
    if sheet not in wb.sheetnames:
        return None
    v = wb[sheet][cell.replace("$", "")].value
    return None if isinstance(v, str) and v.startswith("=") else v


def read_planner(data):
    """(ProjectData, copies, Sizes, warnings, version) from an older planner workbook's bytes. copies: the other
    input tables, {table: (headers, rows)}, for copy_inputs."""
    from .build import ProjectData, Sizes
    from .settings import settings_rows, week_preset
    from .tasks import LAYOUT, N, take_links
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data))
    except Exception as e:                                   # not a zip / not a workbook
        raise UpgradeError(f"This isn't an Excel workbook ({e.__class__.__name__}).")
    tables = _tables(wb)
    if "Tasks" not in tables:
        raise UpgradeError("This workbook has no Tasks table: is it a planner workbook?")
    warnings = []
    version = _named_value(wb, "PlannerVersion") or "1.6 or earlier"

    settings = {}
    for n, _label, _default, _fmt, is_input, _note in settings_rows(True, 26):
        v = _named_value(wb, n)
        if is_input and v is not None:
            settings[n] = v

    if "WorkingWeek" not in settings:                # before v1.7 the week was only the seven Yes / No days
        settings["WorkingWeek"] = week_preset(settings)

    key_of = {N[k]: k for k, kind, *_ in LAYOUT if kind == "in"}
    # before v1.7 each link was three inputs (Pred1 UID / Type / Lag ...): they become the Predecessors text
    key_of.update({f"Pred{k} {label}": f"{x}{k}" for k in range(1, 10) for x, label in (("P", "UID"), ("T", "Type"),
                                                                                       ("L", "Lag"))})
    key_of.update({"Resource": "Res", "Units": "Units"})      # before v1.7: one person + the Assignments sheet
    tasks = []
    for rec in _records(tables, "Tasks", RENAMED["Tasks"]):
        row = {key_of[h]: v for h, v in rec.items() if h in key_of and v not in (None, "")}
        tasks.append(take_links(row))
    while tasks and not tasks[-1].get("Name"):
        tasks.pop()

    resources = [(r.get("Resource Name"), r.get("Initials"), r.get("Group"), r.get("Max Units"), r.get("Std Rate"),
                  r.get("Email"), r.get("Type") or "Work")
                 for r in _records(tables, "ResourceTable") if r.get("Resource Name")]
    holidays = [(r["Date"], r.get("Description")) for r in _records(tables, "HolidayTable") if r.get("Date")]
    leave, months = {}, None
    # leave: a Leave sheet before v1.7, month columns on Resources (Nov-26 ...) since
    source, key = ("LeaveTable", "Resource") if "LeaveTable" in tables else ("ResourceTable", "Resource Name")
    if source in tables:
        months = [h for h in tables[source][1] if isinstance(h, str) and MONTH.match(h)] or None
        for r in _records(tables, source):
            entries = {m: str(r[m]) for m in months or [] if r.get(m) not in (None, "")}
            if r.get(key) and entries:
                leave[r[key]] = entries
    people = {r[0] for r in resources}
    for person in sorted(set(leave) - people):
        warnings.append(f"Leave for {person} was not carried: {person} is not on the Resources sheet.")
        del leave[person]
    assignments = [(r["Task UID"], r.get("Resource"), r.get("Units"))
                   for r in _records(tables, "AssignmentTable") if r.get("Task UID")]

    copies = {name: (head, rows) for name, (_, head, rows) in tables.items() if name not in BUILT}
    default = Sizes()
    sizes = Sizes(tasks=max(default.tasks, len(tasks)), resources=max(default.resources, len(resources)),
                  sprints=max(default.sprints, len(tables.get("SprintTable", (None, None, []))[2])))
    project = ProjectData(settings=settings, tasks=tasks, resources=resources, holidays=holidays, leave=leave,
                          leave_months=months, assignments=assignments)
    return project, copies, sizes, warnings, version


def copy_inputs(wb, copies, warnings):
    """Write the old tables' typed values into the new workbook's input columns with the same table and header.
    Dropdown lists (List... tables) keep the template's values and add the user's extra ones."""
    new = _tables(wb)
    for name, (old_head, old_rows) in copies.items():
        if name not in new:
            if any(v not in (None, "") for row in old_rows for v in row):
                warnings.append(f"The {name} table is not in the new template; its contents were not carried.")
            continue
        ws, head, _ = new[name]
        table = ws.tables[name]
        c1, r1, c2, r2 = range_boundaries(table.ref)
        for i, h in enumerate(head):
            fill = (ws.cell(r1, c1 + i).fill.start_color.rgb or "")[-6:]
            if h not in old_head or fill != INPUT_FILL:
                continue
            values = [row[old_head.index(h)] for row in old_rows]
            if name.startswith("List"):
                current = [ws.cell(r, c1 + i).value for r in range(r1 + 1, r2 + 1)]
                values = [v for v in current if v is not None] + [v for v in values if v not in current and v is not None]
                if len(values) > r2 - r1:                # grow the list's table (lists sit side by side)
                    r2 = r1 + len(values)
                    table.ref = f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}"
                    for r in range(r1 + 1, r2 + 1):
                        ws.cell(r, c1 + i).font, ws.cell(r, c1 + i).border = INPUT_FONT, BOX
            room = r2 - r1
            if any(v not in (None, "") for v in values[room:]):
                warnings.append(f"{name}: only the first {room} rows of '{h}' fit; add rows to the table and copy "
                                "the rest from the old workbook.")
            for k, v in enumerate(values[:room]):
                ws.cell(r1 + 1 + k, c1 + i).value = v


def convert(data):
    """(xlsx bytes of the latest template holding the old workbook's inputs, warnings)."""
    from .build import make_workbook
    project, copies, sizes, warnings, _version = read_planner(data)
    wb = make_workbook(sizes=sizes, data=project)
    copy_inputs(wb, copies, warnings)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), warnings
