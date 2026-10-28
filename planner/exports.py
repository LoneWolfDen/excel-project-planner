"""Exports beside MS Project XML (template 1.2, but they read any planner workbook):

  * B-506 network diagram: the links as a Mermaid flowchart (critical tasks highlighted), which GitHub, Azure
    DevOps, Confluence and mermaid.live draw; or Graphviz DOT.
  * B-706 CSV for other tools: Jira (issue import), Azure DevOps (work-item import) and Microsoft Planner (via its
    Excel / CSV import add-ins), one row per task with the fields each one maps on import.

Both read the workbook's calculated values through planner.mspdi's reader, so the file must have been opened and
saved in Excel (or built with --recalc). Pure standard library, like the rest of the converter code.
"""
import csv
import io
import re

from .mspdi import _date, _links, _link, _num, _text, read_plan


def _tasks(data):
    plan = read_plan(data)
    return plan, [t for t in plan["tasks"] if _text(t.get("Task Name")).strip()]


def _uid(t):
    return int(_num(t.get("UID")))


# ---------------------------------------------------------------- B-506 -----------------------------------------

def _label(text):
    return re.sub(r'["\[\]{}|<>]', "'", str(text))[:60]


def network(data, fmt="mermaid"):
    """The plan's links as a diagram: detail tasks are nodes (milestones round, critical ones red), summaries are
    subgraphs, each link an arrow labelled with its type and lag when not a plain finish-to-start."""
    _plan, tasks = _tasks(data)
    edges = []
    for t in tasks:
        for k in _links(t):
            pred = _num(_link(t, k, "UID"), None)
            if pred and not _text(t.get(f"_Pred{k}Problem")):
                typ, lag = _text(_link(t, k, "Type")) or "FS", _num(_link(t, k, "Lag"))
                label = "" if typ == "FS" and not lag else f"{typ}{'+' if lag > 0 else ''}{lag:g}d" if lag else typ
                edges.append((int(pred), _uid(t), label))
    if fmt == "dot":
        out = ["digraph plan {", "  rankdir=LR;", '  node [shape=box, style="rounded", fontname="Segoe UI"];']
        for t in tasks:
            if not t.get("_IsSummary"):
                crit = t.get("Critical") == "Yes"
                shape = "ellipse" if _num(t.get("_Days", t.get("Duration"))) == 0 else "box"
                out.append(f'  t{_uid(t)} [label="{_uid(t)}: {_label(t["Task Name"])}", shape={shape}'
                           + (', color="#C50F1F", fontcolor="#C50F1F"' if crit else "") + "];")
        out += [f'  t{a} -> t{b}' + (f' [label="{lbl}"]' if lbl else "") + ";" for a, b, lbl in edges]
        return "\n".join(out + ["}"]) + "\n"
    out = ["flowchart LR"]
    depth = 0
    for i, t in enumerate(tasks):                     # summaries open a subgraph; it closes when the level drops
        level = int(_num(t.get("Level"), 1) or 1)
        while depth >= level:
            out.append("  " * depth + "end")
            depth -= 1
        if t.get("_IsSummary"):
            out.append("  " * (depth + 1) + f'subgraph s{_uid(t)}["{_label(t["Task Name"])}"]')
            depth += 1
        else:
            text = f'{_uid(t)}: {_label(t["Task Name"])}'
            milestone = _num(t.get("_Days", t.get("Duration"))) == 0
            out.append("  " * (depth + 1) + f't{_uid(t)}' + (f'(["{text}"])' if milestone else f'["{text}"]'))
    while depth:
        out.append("  " * depth + "end")
        depth -= 1
    out += [f"  t{a} -->" + (f"|{lbl}|" if lbl else "") + f" t{b}" for a, b, lbl in edges]
    critical = [f"t{_uid(t)}" for t in tasks if t.get("Critical") == "Yes" and not t.get("_IsSummary")]
    if critical:
        out += ["  classDef critical fill:#FDE7E9,stroke:#C50F1F,color:#C50F1F", f"  class {','.join(critical)} critical"]
    return "\n".join(out) + "\n"


# ---------------------------------------------------------------- B-706 -----------------------------------------

def _first_person(t):
    return _text(t.get("Resource Names")).split(",")[0].split("[")[0].strip()


def _parent(tasks, i):
    level = _num(tasks[i].get("Level"), 1)
    for j in range(i - 1, -1, -1):
        if _num(tasks[j].get("Level"), 1) < level:
            return tasks[j]
    return None


def _iso(value):
    d = _date(value)
    return d.isoformat() if d else ""


FORMATS = {
    # Jira's CSV importer: map these columns on import (Issue Id / Parent Id keep the hierarchy)
    "jira": ["Issue Id", "Parent Id", "Issue Type", "Summary", "Description", "Assignee", "Start date", "Due date",
             "Story Points", "Sprint", "Labels"],
    # Azure DevOps: Boards > Work items > Import from CSV (ID blank = new; Parent by title is not supported, so the
    # hierarchy goes in Area / tags)
    "azure": ["ID", "Work Item Type", "Title", "Assigned To", "State", "Start Date", "Target Date", "Story Points",
              "Iteration Path", "Description", "Tags"],
    # Microsoft Planner (Planner Premium / Project for the web import, or a Power Automate flow): bucket = phase
    "planner": ["Task Name", "Bucket", "Assigned To", "Start Date", "Due Date", "Progress", "Priority", "Notes"],
}


def _state(t, fmt):
    prog = _num(t.get("Progress"))
    if fmt == "azure":
        return "Done" if prog >= 1 else "Active" if prog > 0 else "New"
    return "Completed" if prog >= 1 else "In progress" if prog > 0 else "Not started"


def to_csv(data, fmt):
    """The plan as CSV text for Jira, Azure DevOps or Microsoft Planner."""
    if fmt not in FORMATS:
        raise ValueError(f"format must be one of {', '.join(FORMATS)}")
    _plan, tasks = _tasks(data)
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(FORMATS[fmt])
    for i, t in enumerate(tasks):
        summary = bool(t.get("_IsSummary"))
        parent = _parent(tasks, i)
        phase = next((p for p in [parent] if p), None)
        while phase is not None and _parent(tasks, tasks.index(phase)) is not None:
            phase = _parent(tasks, tasks.index(phase))
        crit = t.get("Critical") == "Yes"
        start, due = _iso(t.get("Start")), _iso(t.get("Finish"))
        points = _num(t.get("Story Points"), None)
        if fmt == "jira":
            w.writerow([_uid(t), _uid(parent) if parent else "", "Epic" if summary else "Task", t["Task Name"],
                        _text(t.get("Notes")), "" if summary else _first_person(t), start, due,
                        "" if points is None else points, _text(t.get("Sprint")), "critical" if crit else ""])
        elif fmt == "azure":
            w.writerow(["", "Epic" if summary and not parent else "Feature" if summary else "Task", t["Task Name"],
                        "" if summary else _first_person(t), _state(t, fmt), start, due,
                        "" if points is None else points, _text(t.get("Sprint")), _text(t.get("Notes")),
                        "; ".join(x for x in (phase["Task Name"] if phase else "", "critical" if crit else "") if x)])
        elif not summary:                               # Planner has buckets, not summary tasks
            w.writerow([t["Task Name"], phase["Task Name"] if phase else "", _text(t.get("Resource Names")), start,
                        due, _state(t, fmt), "Urgent" if crit else "Medium", _text(t.get("Notes"))])
    return buf.getvalue()
