"""Export a planner workbook to Microsoft Project XML (MSPDI), which MS Project opens directly (File > Open).

Unlike the Import Wizard sheets, MSPDI carries the calendar (working week, hours per day, holidays), links by UID,
constraints, deadlines, actuals, the baseline, resources and assignments, so MS Project schedules the plan the way
the workbook does.

This module uses only the Python standard library, so the same file runs in two places:
  * the command line:  python scripts/export_mspdi.py Project_Planner.xlsx project.xml
  * the browser page docs/export.html, through Pyodide (Python in the browser): nothing to install, works on a phone.
It reads the values Excel saved with the workbook (cached results), so the file must have been saved by Excel (or
LibreOffice) after its last edit, which every normal save does.

Field mapping, in one place for both:
  Tasks row            -> <Task>: UID, ID, Name, WBS, OutlineLevel, Start / Finish (forecast), Duration, Work,
                          Milestone, Summary, Manual, % Complete, actuals, constraint (all eight types),
                          Deadline, Notes, links (Pred UID + type + lag), Baseline, Text1 = Sprint, Number1 = Points
  Settings             -> project name, manager, start, status date, calendar (WorkMon..WorkSun, HoursPerDay)
  HolidayTable         -> calendar exceptions (non-working days)
  ResourceTable        -> <Resource>: name, initials, group, type, max units, standard rate, email
  Tasks Resource+Units -> <Assignment>
Elements are written in the order MS Project writes them (ORDER below); MS Project reports a file with elements out
of order as corrupt.
"""
import datetime as dt
import io
import posixpath
import re
import xml.etree.ElementTree as ET
import zipfile

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
MSPDI_NS = "http://schemas.microsoft.com/project"
EXCEL_EPOCH = dt.date(1899, 12, 30)
LINK_TYPES = {"FF": 0, "FS": 1, "SF": 2, "SS": 3}          # MSPDI PredecessorLink/Type
RESOURCE_TYPES = {"Material": 0, "Work": 1, "Cost": 2}     # MSPDI Resource/Type
TEXT1, NUMBER1 = "188743731", "188743767"                  # MS Project field IDs of Text1 and Number1
DAY_NAMES = ["WorkSun", "WorkMon", "WorkTue", "WorkWed", "WorkThu", "WorkFri", "WorkSat"]   # MSPDI DayType 1..7

# Child order of each element, as MS Project 2016-2019 writes it (taken from MSPDI files saved by MS Project), for
# the elements this export uses. Four tags those files don't contain are placed by their neighbours in Microsoft's
# schema: ActualFinish (after ActualStart), Deadline (after ConstraintDate), Notes and the project's StatusDate.
ORDER = {
    "Project": ["SaveVersion", "Name", "Title", "Manager", "ScheduleFromStart", "StartDate", "CriticalSlackLimit",
                "CalendarUID", "DefaultStartTime", "DefaultFinishTime", "MinutesPerDay", "MinutesPerWeek",
                "DaysPerMonth", "DurationFormat", "WorkFormat", "HonorConstraints", "StatusDate", "ExtendedAttributes",
                "Calendars", "Tasks", "Resources", "Assignments"],
    "Task": ["UID", "ID", "Name", "Active", "Manual", "Type", "IsNull", "WBS", "OutlineNumber", "OutlineLevel",
             "Start", "Finish", "Duration", "DurationFormat", "Work", "Milestone", "Summary", "PercentComplete",
             "ActualStart", "ActualFinish", "ActualCost", "ConstraintType", "ConstraintDate", "Deadline", "Notes",
             "PredecessorLink", "Baseline", "ExtendedAttribute"],
    "PredecessorLink": ["PredecessorUID", "Type", "LinkLag", "LagFormat"],
    "Baseline": ["Number", "Cost", "Duration", "Finish", "Start", "Work", "DurationFormat"],
    "ExtendedAttribute": ["FieldID", "FieldName", "Alias", "Value"],
    "Calendar": ["UID", "Name", "IsBaseCalendar", "WeekDays"],
    "WeekDay": ["DayType", "DayWorking", "TimePeriod", "WorkingTimes"],
    "TimePeriod": ["FromDate", "ToDate"],
    "WorkingTime": ["FromTime", "ToTime"],
    "Resource": ["UID", "ID", "Name", "Type", "IsNull", "Initials", "Group", "EmailAddress", "MaxUnits",
                 "StandardRate", "StandardRateFormat"],
    "Assignment": ["UID", "TaskUID", "ResourceUID", "Units", "Work"],
}


class ExportError(Exception):
    """A problem the user can fix (shown as-is by the command line and the browser page)."""


# ---------------------------------------------------------------- reading the .xlsx (values Excel saved) --------

def _cell_ref(ref):
    m = re.fullmatch(r"\$?([A-Z]+)\$?(\d+)", ref)
    col = 0
    for ch in m.group(1):
        col = col * 26 + ord(ch) - 64
    return int(m.group(2)), col


class Workbook:
    """Just enough of an .xlsx reader: sheets' cached values, defined names and tables."""

    def __init__(self, data):
        try:
            self.zip = zipfile.ZipFile(io.BytesIO(data) if isinstance(data, (bytes, bytearray)) else data)
        except zipfile.BadZipFile:
            raise ExportError("This isn't an .xlsx file. Choose the planner workbook (Project_Planner.xlsx).")
        self.strings = self._shared_strings()
        wb = ET.fromstring(self.zip.read("xl/workbook.xml"))
        rels = self._rels("xl/workbook.xml")
        self.sheet_paths = {s.get("name"): rels[s.get(f"{{{NS['r']}}}id")] for s in wb.find("m:sheets", NS)}
        names = wb.find("m:definedNames", NS)
        self.names = {} if names is None else {n.get("name"): n.text for n in names if n.get("localSheetId") is None}
        self.tables = {}
        for sheet, path in self.sheet_paths.items():
            for target in self._rels(path, "table").values():
                t = ET.fromstring(self.zip.read(target))
                self.tables[t.get("displayName")] = (sheet, t.get("ref"))
        self._values = {}

    def _rels(self, part, kind=None):
        folder, name = posixpath.split(part)
        path = posixpath.join(folder, "_rels", name + ".rels")
        if path not in self.zip.namelist():
            return {}
        out = {}
        for r in ET.fromstring(self.zip.read(path)):
            if kind is None or r.get("Type").endswith("/" + kind):
                target = r.get("Target")       # relative (Excel, LibreOffice) or absolute from the root (openpyxl)
                out[r.get("Id")] = target.lstrip("/") if target.startswith("/") else \
                    posixpath.normpath(posixpath.join(folder, target))
        return out

    def _shared_strings(self):
        if "xl/sharedStrings.xml" not in self.zip.namelist():
            return []
        root = ET.fromstring(self.zip.read("xl/sharedStrings.xml"))
        return ["".join(t.text or "" for t in si.iter(f"{{{NS['m']}}}t")) for si in root.findall("m:si", NS)]

    def values(self, sheet):
        """{(row, col): value} of a sheet's cached values (numbers, text, booleans; errors as text)."""
        if sheet not in self._values:
            out = {}
            root = ET.fromstring(self.zip.read(self.sheet_paths[sheet]))
            for c in root.iter(f"{{{NS['m']}}}c"):
                kind, v = c.get("t"), c.find("m:v", NS)
                if kind == "inlineStr":
                    value = "".join(t.text or "" for t in c.iter(f"{{{NS['m']}}}t"))
                elif v is None or v.text is None:
                    continue
                elif kind == "s":
                    value = self.strings[int(v.text)]
                elif kind == "b":
                    value = v.text == "1"
                elif kind in ("str", "e"):
                    value = v.text
                else:
                    value = float(v.text)
                    value = int(value) if value.is_integer() else value
                out[_cell_ref(c.get("r"))] = value
            self._values[sheet] = out
        return self._values[sheet]

    def named(self, name):
        if name not in self.names:
            raise ExportError(f"The workbook has no '{name}' setting: is this a planner workbook?")
        sheet, ref = self.names[name].rsplit("!", 1)
        return self.values(sheet.strip("'")).get(_cell_ref(ref))

    def table(self, name):
        """Rows of an Excel Table as dicts keyed by header."""
        if name not in self.tables:
            raise ExportError(f"The workbook has no '{name}' table: is this a planner workbook?")
        sheet, ref = self.tables[name]
        (r1, c1), (r2, c2) = (_cell_ref(x) for x in ref.split(":"))
        cells = self.values(sheet)
        head = [cells.get((r1, c)) for c in range(c1, c2 + 1)]
        return [{h: cells.get((r, c)) for h, c in zip(head, range(c1, c2 + 1))} for r in range(r1 + 1, r2 + 1)]


# ---------------------------------------------------------------- the plan, in plain Python values -------------

def _date(value):
    """An Excel serial date (or blank) as a date."""
    if value in (None, "") or isinstance(value, (str, bool)):
        return None
    return EXCEL_EPOCH + dt.timedelta(days=int(value))


def _num(value, default=0):
    return value if isinstance(value, (int, float)) and not isinstance(value, bool) else default


def _text(value):
    return "" if value is None else str(value)


def _working_days(wb):
    """Sun..Sat working flags from the WorkWeek mask (Mon..Sun, 1 = not working), which covers the Working week
    presets; workbooks without it fall back to the Yes / No day settings."""
    mask = _text(wb.named("WorkWeek"))
    if len(mask) == 7 and set(mask) <= {"0", "1"}:
        return [mask[6] == "0"] + [m == "0" for m in mask[:6]]
    return [_text(wb.named(d)) == "Yes" for d in DAY_NAMES]


def read_plan(data):
    """Everything the export needs, from the workbook's cached values."""
    wb = Workbook(data)
    hours = _num(wb.named("HoursPerDay"), 8) or 8
    start = _date(wb.named("ProjectStart"))
    tasks = [t for t in wb.table("Tasks") if _text(t.get("Task Name")).strip()]
    if tasks and tasks[0].get("ID") in (None, ""):
        raise ExportError("The workbook has no calculated values. Open it in Excel, let it calculate, save it, "
                          "and try again.")
    if start is None or start.year < 1950:
        raise ExportError("Settings > Project start date is empty: enter it in Excel, save, and try again.")
    return {
        "name": _text(wb.named("ProjectName")) or "Project",
        "manager": _text(wb.named("ProjectManager")),
        "start": start,
        "status": _date(wb.named("StatusDate")),
        "hours": hours,
        "units": _num(wb.named("DefaultUnits"), 1) or 1,
        "working": _working_days(wb),
        "holidays": sorted({d for d in (_date(h.get("Date")) for h in wb.table("HolidayTable")) if d}),
        "tasks": tasks,
        "resources": [r for r in wb.table("ResourceTable") if _text(r.get("Resource Name")).strip()],
        # more people on tasks (workbooks before v1.5 have no Assignments sheet); _Valid = 1 for the rows Tasks uses
        "assignments": [a for a in (wb.table("AssignmentTable") if "AssignmentTable" in wb.tables else [])
                        if _num(a.get("_Valid")) == 1],
    }


# ---------------------------------------------------------------- writing MSPDI --------------------------------

def _hours_text(h):
    """An MSPDI duration: PT24H0M0S."""
    minutes = round(h * 60)
    return f"PT{minutes // 60}H{minutes % 60}M0S"


def _clock(minutes):
    return f"{minutes // 60:02d}:{minutes % 60:02d}:00"


def _working_times(hours):
    """Morning and afternoon shifts adding up to hours: 08:00-12:00 and 13:00-17:00 for an 8-hour day."""
    total = round(hours * 60)
    morning = min(total, 240)
    shifts = [(480, 480 + morning)]
    if total > morning:
        shifts.append((780, 780 + total - morning))
    return shifts


def _sub(parent, tag, text=None):
    e = ET.SubElement(parent, tag)
    if text is not None:
        e.text = str(text)
    return e


def _when(day, minutes):
    return f"{day.isoformat()}T{_clock(minutes)}"


def to_mspdi(plan):
    """MSPDI XML text for a plan from read_plan."""
    hours = plan["hours"]
    shifts = _working_times(hours)
    day_start, day_end = shifts[0][0], shifts[-1][1]
    start_of = lambda d: _when(d, day_start)
    end_of = lambda d: _when(d, day_end)
    flag = lambda b: "1" if b else "0"

    p = ET.Element("Project")
    p.set("xmlns", MSPDI_NS)
    _sub(p, "SaveVersion", 14)
    _sub(p, "Name", plan["name"] + ".xml")
    _sub(p, "Title", plan["name"])
    if plan["manager"]:
        _sub(p, "Manager", plan["manager"])
    _sub(p, "ScheduleFromStart", 1)
    _sub(p, "StartDate", start_of(plan["start"]))
    _sub(p, "CriticalSlackLimit", 0)
    _sub(p, "CalendarUID", 1)
    _sub(p, "DefaultStartTime", _clock(day_start))
    _sub(p, "DefaultFinishTime", _clock(day_end))
    _sub(p, "MinutesPerDay", round(hours * 60))
    _sub(p, "MinutesPerWeek", round(hours * 60 * max(1, sum(plan["working"]))))
    _sub(p, "DaysPerMonth", 20)
    _sub(p, "DurationFormat", 7)             # days
    _sub(p, "WorkFormat", 2)                 # hours
    _sub(p, "HonorConstraints", 1)
    if plan["status"]:
        _sub(p, "StatusDate", end_of(plan["status"]))
    attrs = _sub(p, "ExtendedAttributes")
    for field, alias in ((TEXT1, "Sprint"), (NUMBER1, "Story Points")):
        a = _sub(attrs, "ExtendedAttribute")
        _sub(a, "FieldID", field)
        _sub(a, "FieldName", "Text1" if field == TEXT1 else "Number1")
        _sub(a, "Alias", alias)

    cal = _sub(_sub(p, "Calendars"), "Calendar")
    _sub(cal, "UID", 1)
    _sub(cal, "Name", "Standard")
    _sub(cal, "IsBaseCalendar", 1)
    week = _sub(cal, "WeekDays")
    for day_type, working in enumerate(plan["working"], start=1):
        wd = _sub(week, "WeekDay")
        _sub(wd, "DayType", day_type)
        _sub(wd, "DayWorking", flag(working))
        if working:
            times = _sub(wd, "WorkingTimes")
            for a, b in shifts:
                t = _sub(times, "WorkingTime")
                _sub(t, "FromTime", _clock(a))
                _sub(t, "ToTime", _clock(b))
    for day in plan["holidays"]:                 # exceptions: DayType 0 with a date range, not working
        wd = _sub(week, "WeekDay")
        _sub(wd, "DayType", 0)
        _sub(wd, "DayWorking", 0)
        period = _sub(wd, "TimePeriod")
        _sub(period, "FromDate", _when(day, 0))
        _sub(period, "ToDate", f"{day.isoformat()}T23:59:00")

    resources = {r["Resource Name"]: i for i, r in enumerate(plan["resources"], start=1)}
    tasks_el = _sub(p, "Tasks")
    assignments = []
    for t in plan["tasks"]:
        summary = bool(t.get("_IsSummary"))
        days = _num(t.get("Duration"))
        if summary:
            days = _num(t.get("BL Source Duration"))
        milestone = not summary and days == 0
        manual = t.get("Mode") == "Manual"
        cons = _date(t.get("Constraint Date", t.get("Start No Earlier Than")))     # (the column's name before v1.5)
        code = int(_num(t.get("_ConstraintCode"), 4 if cons else 0))
        fs, ff = _date(t.get("Forecast Start")), _date(t.get("Forecast Finish"))
        if manual and cons:
            fs = _date(t.get("_ConstraintStart")) or cons
        work = 0 if summary else _num(t.get("Work"))
        e = _sub(tasks_el, "Task")
        _sub(e, "UID", int(t["UID"]))
        _sub(e, "ID", int(_num(t.get("ID"))))
        _sub(e, "Name", t["Task Name"])
        _sub(e, "Type", 0)                      # fixed units, MS Project's default
        _sub(e, "IsNull", 0)
        _sub(e, "WBS", _text(t.get("WBS")))
        _sub(e, "OutlineNumber", _text(t.get("WBS")))
        _sub(e, "OutlineLevel", int(_num(t.get("Level"), 1)) or 1)
        # a milestone sits at the end of its predecessor's day (MS Project's FS + 0), otherwise at the day's start
        at_end = milestone and _has_links(t)
        if fs:
            _sub(e, "Start", end_of(fs) if at_end else start_of(fs))
        if ff:
            _sub(e, "Finish", end_of(ff) if at_end or not milestone else start_of(ff))
        _sub(e, "Duration", _hours_text(days * hours))
        _sub(e, "DurationFormat", 7)
        _sub(e, "Work", _hours_text(work))
        _sub(e, "Milestone", flag(milestone))
        _sub(e, "Summary", flag(summary))
        _sub(e, "PercentComplete", round(_num(t.get("Progress")) * 100))
        actual_start, actual_finish = _date(t.get("Actual Start")), _date(t.get("Actual Finish"))
        if actual_start and not summary:
            _sub(e, "ActualStart", start_of(actual_start))
        if actual_finish and not summary:
            _sub(e, "ActualFinish", end_of(actual_finish))
        if not summary and _num(t.get("Actual Cost"), None) is not None:
            _sub(e, "ActualCost", t["Actual Cost"])
        # ConstraintType is MS Project's code (the planner's _ConstraintCode); codes 2-7 carry a date, which for
        # Must Finish On / Finish No Earlier / Later Than (3, 6, 7) is the end of that day
        if summary or manual:
            code = 0
        _sub(e, "ConstraintType", code)
        if code >= 2 and cons:
            _sub(e, "ConstraintDate", end_of(cons) if code in (3, 6, 7) else start_of(cons))
        deadline = _date(t.get("Deadline"))
        if deadline:
            _sub(e, "Deadline", end_of(deadline))
        if _text(t.get("Notes")):
            _sub(e, "Notes", t["Notes"])
        if not summary:
            for k in _links(t):
                pred = _num(t.get(f"Pred{k} UID"))
                if pred and not _text(t.get(f"_Pred{k}Problem")):
                    link = _sub(e, "PredecessorLink")
                    _sub(link, "PredecessorUID", int(pred))
                    _sub(link, "Type", LINK_TYPES.get(_text(t.get(f"Pred{k} Type")) or "FS", 1))
                    _sub(link, "LinkLag", round(_num(t.get(f"Pred{k} Lag")) * hours * 600))  # tenths of a minute
                    _sub(link, "LagFormat", 7)
        _sub(e, "Active", 1)
        _sub(e, "Manual", flag(manual))
        if not summary:
            for field, value in ((TEXT1, _text(t.get("Sprint"))), (NUMBER1, _num(t.get("Story Points"), None))):
                if value not in ("", None):
                    a = _sub(e, "ExtendedAttribute")
                    _sub(a, "FieldID", field)
                    _sub(a, "Value", value)
        base_start, base_finish = _date(t.get("Baseline Start")), _date(t.get("Baseline Finish"))
        if base_start and base_finish:
            b = _sub(e, "Baseline")
            _sub(b, "Number", 0)
            _sub(b, "Start", start_of(base_start))
            _sub(b, "Finish", end_of(base_finish))
            _sub(b, "Duration", _hours_text(_num(t.get("Baseline Duration")) * hours))
            _sub(b, "DurationFormat", 7)
            _sub(b, "Work", _hours_text(_num(t.get("Baseline Work"))))
            _sub(b, "Cost", _num(t.get("Baseline Cost")))
        name = _text(t.get("Resource"))
        if name and not summary:
            if name not in resources:
                raise ExportError(f"Task {t['UID']} ({t['Task Name']}) uses resource '{name}', which isn't on the "
                                  "Resources sheet.")
            units = _num(t.get("Units"), None)
            assignments.append((int(t["UID"]), resources[name], plan["units"] if units is None else units,
                                _num(t.get("_LeafWork"), work)))
    for a in plan["assignments"]:                # checked by the sheet: a detail task, a resource on Resources
        units = _num(a.get("Units"), None)
        assignments.append((int(a["Task UID"]), resources[a["Resource"]], plan["units"] if units is None else units,
                            _num(a.get("Work"))))

    res_el = _sub(p, "Resources")
    for uid, r in enumerate(plan["resources"], start=1):
        e = _sub(res_el, "Resource")
        _sub(e, "UID", uid)
        _sub(e, "ID", uid)
        _sub(e, "Name", r["Resource Name"])
        _sub(e, "Type", RESOURCE_TYPES.get(_text(r.get("Type")), 1))
        _sub(e, "IsNull", 0)
        for tag, key in (("Initials", "Initials"), ("Group", "Group"), ("EmailAddress", "Email")):
            if _text(r.get(key)):
                _sub(e, tag, r[key])
        _sub(e, "MaxUnits", _num(r.get("Max Units"), 1))
        _sub(e, "StandardRate", _num(r.get("Std Rate")))
        _sub(e, "StandardRateFormat", 2)          # per hour

    asg_el = _sub(p, "Assignments")
    for uid, (task, resource, units, work) in enumerate(assignments, start=1):
        e = _sub(asg_el, "Assignment")
        _sub(e, "UID", uid)
        _sub(e, "TaskUID", task)
        _sub(e, "ResourceUID", resource)
        _sub(e, "Units", units)
        _sub(e, "Work", _hours_text(work))

    _put_in_order(p)
    ET.indent(p)
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + ET.tostring(p, encoding="unicode") + "\n"


def _put_in_order(element):
    """Sort every element's children into ORDER (stable, so repeated tags keep their sequence)."""
    order = ORDER.get(element.tag)
    if order:
        element[:] = sorted(element, key=lambda c: order.index(c.tag))
    for child in element:
        _put_in_order(child)


def _links(t):
    """The link numbers the workbook has (Pred1 ... PredN): read from its headers, so any LINKS setting works."""
    k = 1
    while f"Pred{k} UID" in t:
        yield k
        k += 1


def _has_links(t):
    return any(_num(t.get(f"Pred{k} UID")) and not _text(t.get(f"_Pred{k}Problem")) for k in _links(t))


def convert(data):
    """Workbook bytes (or a path) in, MSPDI XML text out: the one call the CLI and the browser page make."""
    return to_mspdi(read_plan(data))
