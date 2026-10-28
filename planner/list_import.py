"""Start a plan from a list (B-914): a pasted outline or a CSV file in, a planner workbook out.

Three inputs, told apart by their content:
  * an outline: one task per line; indentation (a tab or two spaces), "-" / "*" / "•" bullets or "1.2.3" numbers
    give the level; optional fields after ";": a duration ("5d", "5", "0d" or "milestone"), "after 3" or
    "after 3, 5SS+2d" (line numbers, which become the UIDs), and the rest are resources ("Alex Chen, JL[50%]");
  * a CSV with this planner's column names (Level, Task Name, Duration, Predecessors, Resources, Notes, Actual Start,
    Actual Finish, % Complete, Deadline, Story Points);
  * a CSV exported by MS Project (Name, Outline Level, Duration as "5 days", Predecessors by ID as "3FS+2 days",
    Resource Names, ...).
Nothing else is guessed: dates are scheduled by the workbook from durations and links. The people named become the
Resources rows. The workbook is built by planner.build like an MS Project import, so it needs nothing else.

Pure standard library (csv, re), so the same code runs in the CLI (scripts/import_csv.py) and the browser page.
"""
import csv
import datetime as dt
import io
import re

BULLET = re.compile(r"^[-*•·]\s*")
NUMBERED = re.compile(r"^(\d+(?:\.\d+)*)[.)]?\s+")
DURATION = re.compile(r"^(\d+(?:\.\d+)?)\s*(d|day|days|w|wk|wks|week|weeks|h|hr|hrs|hour|hours|mo|mon|month|months)?\??$",
                      re.I)
LINK = re.compile(r"^(\d+)(FS|SS|FF|SF)?([+-]\d+(?:\.\d+)?)?$", re.I)
PLANNER_COLUMNS = {"Level": "Level", "Task Name": "Name", "Duration": "Dur", "Predecessors": "PredIn",
                   "Resources": "ResIn", "Notes": "Notes", "Actual Start": "AS", "Actual Finish": "AF",
                   "% Complete": "Pct", "Deadline": "DL", "Story Points": "SP", "UID": "UID", "Owner": "Owner"}
MSP_COLUMNS = {"Name": "Name", "Outline Level": "Level", "Duration": "Dur", "Predecessors": "PredIn",
               "Resource Names": "ResIn", "Notes": "Notes", "Actual Start": "AS", "Actual Finish": "AF",
               "% Complete": "Pct", "Deadline": "DL", "ID": "ID", "Unique ID": "UID"}


class ListImportError(Exception):
    """Text the importer can't read (shown as-is by the CLI and the page)."""


def days(text, hours_per_day=8):
    """Working days from "5", "5d", "5 days", "1 wk", "16 hrs", "1 mo" (20 days); None when it isn't a duration."""
    m = DURATION.match(str(text).strip().replace("edays", "days"))
    if not m:
        return None
    n, unit = float(m.group(1)), (m.group(2) or "d").lower()
    n = n * 5 if unit.startswith("w") else n / hours_per_day if unit.startswith("h") else \
        n * 20 if unit.startswith("mo") else n
    return int(n) if n == int(n) else round(n, 2)


def links(text, uid_of=lambda n: n):
    """The planner's Predecessors text from "3, 5SS+2d" or MS Project's "3FS+2 days;5" (by ID: uid_of maps)."""
    out, problems = [], []
    for item in re.split(r"[,;]", str(text or "")):
        item = item.strip()
        if not item:
            continue
        compact = re.sub(r"(?i)(e?days?|d)$", "", re.sub(r"\s+", "", item))   # "3FS+2 days" -> "3FS+2"
        m = LINK.match(compact)
        if not m:
            problems.append(item)
            continue
        uid = uid_of(int(m.group(1)))
        kind = (m.group(2) or "FS").upper()
        lag = m.group(3).replace(" ", "") if m.group(3) else ""
        out.append(f"{uid}{'' if kind == 'FS' and not lag else kind}{lag + 'd' if lag else ''}")
    return ", ".join(out), problems


def _initials(name):
    return "".join(w[0] for w in name.split() if w[:1].isalpha()).upper()[:3]


def _people(rows):
    """Resources rows for everyone named in the tasks' Resources text (names before any [units])."""
    names = []
    for row in rows:
        for part in str(row.get("ResIn") or "").split(","):
            who = part.split("[")[0].strip()
            if who and who not in names:
                names.append(who)
    full = [n for n in names if " " in n or n not in {_initials(m) for m in names if m != n}]   # "SR" = Sam Rivera
    return [(n, _initials(n), "", 1, None, "") for n in full]


def from_outline(text):
    """(task rows, warnings) from an indented outline."""
    rows, warnings = [], []
    unit = None
    for number, line in enumerate(l for l in str(text).replace("\r", "").split("\n") if l.strip()):
        raw = line.expandtabs(2)
        indent = len(raw) - len(raw.lstrip())
        body = raw.strip()
        m = NUMBERED.match(body)
        if m:
            level = m.group(1).count(".") + 1
            body = body[m.end():]
        else:
            body = BULLET.sub("", body)
            if indent and unit is None:
                unit = indent                           # the first indent sets the step (2 or 4 spaces, a tab)
            level = 1 + (indent // unit if unit else 0)
        name, *fields = [f.strip() for f in body.split(";")]
        row = {"UID": number + 1, "Level": level, "Name": name}
        for f in fields:
            if not f:
                continue
            if f.lower() in ("milestone", "ms"):
                row["Dur"] = 0
            elif days(f) is not None:
                row["Dur"] = days(f)
            elif f.lower().startswith("after "):
                row["PredIn"], bad = links(f[6:])
                warnings += [f"Line {number + 1} ({name}): can't read the link '{b}'." for b in bad]
            else:
                row["ResIn"] = f
        rows.append(row)
    return rows, warnings


def _date(text):
    text = str(text or "").strip()
    if not text or text.upper() == "NA":
        return None
    text = re.sub(r"^[A-Za-z]{3}\s+(?=\d)", "", text)    # MS Project: "Mon 07/09/26"
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d/%m/%y", "%m/%d/%y", "%d-%b-%Y", "%d-%b-%y", "%d %B %Y"):
        try:
            return dt.datetime.strptime(text.split(" ")[0] if " " not in fmt else text, fmt).date()
        except ValueError:
            continue
    return None


def from_csv(text, hours_per_day=8):
    """(task rows, warnings, the earliest Start in it or None) from a CSV: this planner's columns, or an MS Project
    export."""
    sample = str(text)[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
    except csv.Error:
        dialect = csv.excel
    reader = csv.DictReader(io.StringIO(str(text).lstrip("﻿")), dialect=dialect)
    head = [h.strip() for h in (reader.fieldnames or [])]
    msp = "Name" in head and "Task Name" not in head
    columns = MSP_COLUMNS if msp else PLANNER_COLUMNS
    if not {"Name", "Task Name"} & set(head):
        raise ListImportError("The CSV needs a 'Task Name' column (this planner's headers) or a 'Name' column "
                              "(an MS Project export).")
    rows, warnings, ids, starts = [], [], {}, []
    for i, rec in enumerate(reader):
        rec = {k.strip(): (v or "").strip() for k, v in rec.items() if k}
        get = {key: rec.get(h, "") for h, key in columns.items()}
        if not get["Name"]:
            continue
        uid = int(get["UID"]) if str(get.get("UID", "")).isdigit() else len(rows) + 1
        ids[int(get["ID"]) if str(get.get("ID", "")).isdigit() else len(rows) + 1] = uid
        row = {"UID": uid, "Level": int(float(get["Level"])) if get.get("Level") else 1, "Name": get["Name"]}
        if get.get("Dur"):
            d = days(get["Dur"], hours_per_day)
            if d is None:
                warnings.append(f"Row {i + 2} ({row['Name']}): can't read the duration '{get['Dur']}'.")
            else:
                row["Dur"] = d
        for key in ("Notes", "ResIn", "Owner"):
            if get.get(key):
                row[key] = ", ".join(p.strip() for p in get[key].split(";")) if key == "ResIn" else get[key]
        for key in ("AS", "AF", "DL"):
            if get.get(key) and _date(get[key]):
                row[key] = _date(get[key])
        if _date(rec.get("Start")):
            starts.append(_date(rec["Start"]))
        if get.get("Pct"):
            pct = float(get["Pct"].rstrip("%")) / (100 if get["Pct"].endswith("%") or float(
                get["Pct"].rstrip("%")) > 1 else 1)
            if pct:
                row["Pct"] = pct
        if get.get("SP"):
            row["SP"] = float(get["SP"])
        row["_pred"] = get.get("PredIn", "")
        rows.append(row)
    for row in rows:                                    # MS Project's predecessors are IDs: map them to UIDs
        text = row.pop("_pred")
        if text:
            row["PredIn"], bad = links(text, (lambda n: ids.get(n, n)) if msp else (lambda n: n))
            warnings += [f"{row['Name']}: can't read the link '{b}'." for b in bad]
    for row in rows:                                   # a summary row's Duration is a roll-up: leave it out
        nxt = rows[rows.index(row) + 1] if rows.index(row) + 1 < len(rows) else None
        if nxt and nxt["Level"] > row["Level"]:
            row.pop("Dur", None)
    return rows, warnings, min(starts) if starts else None


def read(text, name="list"):
    """(ProjectData, warnings): a CSV when the file name or the first line says so, otherwise an outline."""
    from .build import ProjectData
    if isinstance(text, bytes):
        for enc in ("utf-8-sig", "cp1252"):
            try:
                text = text.decode(enc)
                break
            except UnicodeDecodeError:
                continue
    first = str(text).lstrip("﻿").split("\n", 1)[0]
    cells = {c.strip().strip('"') for c in re.split(r"[,;\t]", first)}
    is_csv = str(name).lower().endswith(".csv") or bool(cells & {"Task Name", "Name"}) and len(cells) > 1
    settings = {}
    if is_csv:
        rows, warnings, start = from_csv(text)
        if start:                                      # the plan starts where the exported one did
            settings["ProjectStartDate"] = start
    else:
        rows, warnings = from_outline(text)
    if not rows:
        raise ListImportError("There are no tasks in it: one task per line, or a CSV with a Task Name column.")
    position = {r["UID"]: i for i, r in enumerate(rows)}
    below = sum(1 for i, r in enumerate(rows) for uid in re.findall(r"(?:^|,\s*)(\d+)", r.get("PredIn") or "")
                if position.get(int(uid), -1) > i)
    if below:
        warnings.append(f"{below} links point to a task listed below its successor: Link Check shows them; the "
                        "converter page's Reorder option puts them right.")
    return ProjectData(settings=settings, tasks=rows, resources=_people(rows)), warnings


def convert(text, name="list", project_type=None):
    """(xlsx bytes, warnings) for the CLI and the page."""
    from .build import Sizes, make_workbook
    project, warnings = read(text, name)
    if project_type:
        project.settings["ProjectType"] = project_type
    sizes = Sizes(tasks=max(200, len(project.tasks) + 50), resources=max(50, len(project.resources) + 10))
    wb = make_workbook(sizes=sizes, data=project)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), warnings
