"""Import a Microsoft Project XML file (MSPDI: MS Project > File > Save As > XML) into a new planner workbook.

The reverse of planner/mspdi.py. Reading the XML uses only the standard library; the workbook is then built by
planner.build.make_workbook with the project's own data instead of the sample, so it has every sheet, formula and
report of the template. The same code runs on the command line (scripts/import_mspdi.py) and in the browser page
docs/export.html (Pyodide, with openpyxl bundled into the page).

What the planner can't hold is reported (warnings) and, for a task, noted in its Notes:
  * more than four predecessors (tasks.LINKS; the first four are kept), links on summary tasks, lags in % or
    elapsed time;
  * constraints on summary tasks (summaries roll up);
  * part-day calendar exceptions.
A predecessor listed below its successor is kept: Link Check on Tasks explains it (move the task above).
"""
import datetime as dt
import io
import re
import xml.etree.ElementTree as ET

from .settings import week_preset
from .tasks import CONSTRAINTS, LINK_NUMBERS, LINKS, PEOPLE, each, take_links

NS = "{http://schemas.microsoft.com/project}"
LINK_TYPES = {0: "FF", 1: "FS", 2: "SF", 3: "SS"}
RESOURCE_TYPES = {0: "Material", 1: "Work", 2: "Cost"}
DAY_SETTINGS = ["WorkSun", "WorkMon", "WorkTue", "WorkWed", "WorkThu", "WorkFri", "WorkSat"]   # DayType 1..7
NUMBER1 = "188743767"
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
# LagFormat: 7 = working days; 3/5/9/11 = minutes / hours / weeks / months (working time); the e-forms are elapsed
WORKING_LAG = {3, 5, 7, 9, 11, 35, 37, 39, 41, 43}
PERCENT_LAG = {19, 20, 51, 52}


class MSPImportError(Exception):
    """A problem with the file the user can fix (shown as-is by the command line and the browser page)."""


def _text(e, tag, default=None):
    found = e.find(NS + tag)
    return default if found is None or found.text is None else found.text


def _int(e, tag, default=0):
    value = _text(e, tag)
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return default


def _float(e, tag, default=0.0):
    value = _text(e, tag)
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _date(value):
    return dt.date.fromisoformat(value[:10]) if value else None


def _hours(duration):
    """An MSPDI duration such as PT64H0M0S (or P2DT4H0M0S) in hours."""
    if not duration:
        return 0.0
    m = re.fullmatch(r"-?P(?:(\d+(?:\.\d+)?)D)?(?:T(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?)?",
                     duration)
    if not m:
        return 0.0
    d, h, mi, s = (float(x) if x else 0.0 for x in m.groups())
    return d * 24 + h + mi / 60 + s / 3600


def _days(hours, hours_per_day):
    return round(hours / hours_per_day, 2) if hours_per_day else 0


def _calendar_days(calendar):
    """(working week flags Sun..Sat or None, [non-working dates], [part-day dates]) from a Calendar element."""
    week, off, partial = None, [], []
    weekdays = calendar.find(NS + "WeekDays")
    for wd in [] if weekdays is None else weekdays:
        kind, working = _int(wd, "DayType"), _text(wd, "DayWorking") == "1"
        if 1 <= kind <= 7:
            week = week or [None] * 7
            week[kind - 1] = working
        elif kind == 0:                                   # exception (MS Project 2003 form): a date range
            period = wd.find(NS + "TimePeriod")
            if period is not None:
                (off if not working else partial).extend(_range(period))
    exceptions = calendar.find(NS + "Exceptions")
    for ex in [] if exceptions is None else exceptions:   # exceptions (MS Project 2007+ form)
        period = ex.find(NS + "TimePeriod")
        if period is None:
            continue
        working = _text(ex, "DayWorking") == "1"
        (off if not working else partial).extend(_range(period))
    return week, off, partial


def _range(period):
    start, end = _date(_text(period, "FromDate")), _date(_text(period, "ToDate"))
    if not start or not end or (end - start).days > 366:
        return []
    return [start + dt.timedelta(days=k) for k in range((end - start).days + 1)]


def read_project(data):
    """(ProjectData, warnings) from MSPDI XML bytes or text."""
    from .build import ProjectData          # imported here so reading needs only the standard library
    try:
        root = ET.fromstring(data if isinstance(data, (bytes, str)) else data.read())
    except ET.ParseError as e:
        raise MSPImportError(f"This isn't an XML file ({e}). In MS Project use File > Save As > XML.")
    if root.tag != NS + "Project":
        raise MSPImportError("This XML isn't a Microsoft Project file. In MS Project use File > Save As > XML.")
    warnings = []
    hours = (_int(root, "MinutesPerDay", 480) or 480) / 60
    start = _date(_text(root, "StartDate"))
    settings = {"ProjectName": _text(root, "Title") or re.sub(r"\.xml$", "", _text(root, "Name", "Project")),
                "ProjectManager": _text(root, "Manager"), "ProjectStartDate": start,
                "StatusDateInput": _date(_text(root, "StatusDate")), "HoursPerDay": hours}

    calendars = {_int(c, "UID"): c for c in root.iter(NS + "Calendar")}
    base = calendars.get(_int(root, "CalendarUID", 1))
    holidays = []
    if base is not None:
        week, off, partial = _calendar_days(base)
        for name, working in zip(DAY_SETTINGS, week or []):
            if working is not None:
                settings[name] = "Yes" if working else "No"
        settings["WorkingWeek"] = week_preset(settings)
        holidays = [(d, "From MS Project") for d in sorted(set(off))]
        if partial:
            warnings.append(f"{len(partial)} part-day calendar exceptions were not imported (whole days only).")

    resources, by_uid, leave = [], {}, {}
    for r in root.iter(NS + "Resource"):
        name = _text(r, "Name")
        if not name or _int(r, "UID") == 0 or _text(r, "IsNull") == "1":
            continue
        by_uid[_int(r, "UID")] = name
        resources.append((name, _text(r, "Initials"), _text(r, "Group"), _float(r, "MaxUnits", 1.0),
                          round(_float(r, "StandardRate") * hours, 2), _text(r, "EmailAddress"),   # a day rate
                          RESOURCE_TYPES.get(_int(r, "Type", 1), "Work")))
        own = calendars.get(_int(r, "CalendarUID", -1))
        if own is not None and own is not base:
            _, off, _ = _calendar_days(own)
            days = [d for d in off if d.weekday() < 5 and d not in {h for h, _ in holidays}]
            for d in sorted(set(days)):
                label = f"{MONTHS[d.month - 1]}-{d.year % 100:02d}"
                month = leave.setdefault(name, {})
                month[label] = f"{month[label]}, {d.day}" if label in month else str(d.day)

    assigned = {}
    for a in root.iter(NS + "Assignment"):
        task, resource = _int(a, "TaskUID"), _int(a, "ResourceUID")
        if resource in by_uid:
            assigned.setdefault(task, []).append((_hours(_text(a, "Work")), by_uid[resource], _float(a, "Units", 1.0)))

    tasks, outline = [], []
    for t in sorted(root.iter(NS + "Task"), key=lambda t: _int(t, "ID")):
        uid = _int(t, "UID")
        if uid == 0 or _text(t, "IsNull") == "1" or not _text(t, "Name"):
            continue
        tasks.append(_task(t, uid, hours, assigned.get(uid, []), warnings))
        outline.append(tasks[-1]["Level"])
    if not tasks:
        raise MSPImportError("The file has no tasks.")
    summaries = {i for i in range(len(outline) - 1) if outline[i + 1] > outline[i]}
    for i in summaries:
        row = tasks[i]
        if any(row.get(f"P{k}") for k in LINK_NUMBERS):
            _warn(row, warnings, "links on a summary task are ignored (link one of its tasks instead)")
        if row.get("CType") and row["Mode"] == "Auto":
            _warn(row, warnings, f"constraint '{row['CType']}' on a summary task is ignored")
        for key in ("Dur", "AS", "AF", "Pct", "Res", "Units", "SP", "CType", "Cons", "_more"):
            row.pop(key, None)                            # summaries roll up in the planner
    more = [a for row in tasks for a in row.pop("_more", [])]
    position = {row["UID"]: i for i, row in enumerate(tasks)}
    below = sum(1 for i, row in enumerate(tasks) for p in each("P") if position.get(row.get(p), -1) > i)
    if below:
        warnings.append(f"{below} links point to a task listed below its successor; Link Check on Tasks shows them "
                        "(move the predecessor above).")

    for row in tasks:
        take_links(row)                                   # P1 / T1 / L1 ... -> the Predecessors text
    # MS Project works Actual Cost out from rates, as the Cost sheet does; only a fixed cost is typed in
    typed = sum(1 for t in root.iter(NS + "Task") if _float(t, "FixedCost"))
    if typed:
        warnings.append(f"{typed} tasks had fixed costs in MS Project. The planner works cost out on the Cost sheet "
                        "(each person's work x their day rate = MS Project's Standard Rate x hours per day), so fixed "
                        "costs were not imported.")
    months = _leave_months(start, leave)
    return ProjectData(settings=settings, tasks=tasks, resources=resources, holidays=holidays, leave=leave,
                       leave_months=months, assignments=more), warnings


def _warn(row, warnings, message):
    warnings.append(f"Task {row['UID']} ({row['Name']}): {message}.")
    row["Notes"] = ((row.get("Notes") or "") + f" [From MS Project: {message}.]").strip()


def _task(t, uid, hours, assignments, warnings):
    row = {"UID": uid, "Level": max(1, _int(t, "OutlineLevel", 1)), "Name": _text(t, "Name"),
           "Mode": "Manual" if _text(t, "Manual") == "1" else "Auto",
           "Dur": 0 if _text(t, "Milestone") == "1" else _days(_hours(_text(t, "Duration")), hours),
           "Notes": _text(t, "Notes"), "Owner": _text(t, "Contact") or None}
    constraint, when = _int(t, "ConstraintType"), _date(_text(t, "ConstraintDate"))
    if row["Mode"] == "Manual":
        row["Cons"] = _date(_text(t, "Start"))           # manual tasks are pinned to the Constraint Date
    elif constraint == 1 or (2 <= constraint <= 7 and when):
        row["CType"] = CONSTRAINTS[constraint]           # ConstraintType is the index in CONSTRAINTS
        row["Cons"] = when if constraint >= 2 else None
    row["DL"] = _date(_text(t, "Deadline"))
    row["AS"], row["AF"] = _date(_text(t, "ActualStart")), _date(_text(t, "ActualFinish"))
    pct = _float(t, "PercentComplete")
    row["Pct"] = pct / 100 if pct else None

    links = t.findall(NS + "PredecessorLink")
    for k, link in enumerate(links[:LINKS], start=1):
        row[f"P{k}"] = _int(link, "PredecessorUID")
        row[f"T{k}"] = LINK_TYPES.get(_int(link, "Type", 1), "FS")
        lag, fmt = _float(link, "LinkLag"), _int(link, "LagFormat", 7)
        if lag and fmt in PERCENT_LAG:
            _warn(row, warnings, f"a lag in % on the link from UID {row[f'P{k}']} was dropped")
        elif lag:
            days = round(lag / 600 / hours) if fmt in WORKING_LAG else round(lag / 600 / 24 * 5 / 7)
            row[f"L{k}"] = days or None
            if fmt not in WORKING_LAG:
                _warn(row, warnings, f"an elapsed lag from UID {row[f'P{k}']} became {days} working days")
    if len(links) > LINKS:
        dropped = ", ".join(_text(link, "PredecessorUID") for link in links[LINKS:])
        _warn(row, warnings, f"only {LINKS} predecessors are kept; dropped UID {dropped}")

    if assignments:            # into the task's Resources cell, the person with the most work first
        ranked = sorted(assignments, key=lambda a: -a[0])
        _, name, units = ranked[0]
        row["Res"], row["Units"] = name, None if units == 1 else units
        row["_more"] = [(uid, n, None if u == 1 else u) for _, n, u in ranked[1:PEOPLE]]
        if len(ranked) > PEOPLE:
            _warn(row, warnings, f"only {PEOPLE} people per task are kept; dropped "
                                 + ", ".join(n for _, n, _ in ranked[PEOPLE:]))
    for a in t.findall(NS + "ExtendedAttribute"):
        if _text(a, "FieldID") == NUMBER1:
            row["SP"] = _float(a, "Value") or None
    base = next((b for b in t.findall(NS + "Baseline") if _int(b, "Number") == 0), None)
    if base is not None and _text(base, "Start"):
        row["BS"], row["BF"] = _date(_text(base, "Start")), _date(_text(base, "Finish"))
        row["BD"] = _days(_hours(_text(base, "Duration")), hours)
        row["BW"] = round(_hours(_text(base, "Work")), 2)        # the Cost sheet prices it at the day rates
    return row


def _leave_months(start, leave):
    """Leave month headers (Resources sheet): from the project's first month for three years, and any month that has leave."""
    from .leave import month_labels
    first = (start.year, start.month) if start else (2026, 1)
    labels = month_labels(first, 36)
    for months in leave.values():
        for label in months:
            if label not in labels:
                labels.append(label)
    key = lambda label: (int(label[-2:]), MONTHS.index(label[:3]))
    return sorted(labels, key=key)


def convert(data):
    """MSPDI XML in, (planner workbook .xlsx bytes, warnings) out: the one call the CLI and the browser page make."""
    from .build import Sizes, make_workbook
    project, warnings = read_project(data)
    n = len(project.tasks)
    people = len(project.resources)
    sizes = Sizes(tasks=max(200, n + 50), resources=max(50, people + 10), grid_resources=max(25, min(people, 500)))
    wb = make_workbook(sizes=sizes, data=project)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), warnings
