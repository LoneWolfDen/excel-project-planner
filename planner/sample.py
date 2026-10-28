"""Sample project, and a generator that repeats it to make large stress-test workbooks."""
import datetime as dt

import re

from .tasks import LINKS, PEOPLE, take_links

D = dt.date

# Level, Task Name, Mode, Duration, Pred1 UID, Type, Lag, Pred2 UID, Type, Lag, Constraint Date, Deadline,
# Actual Start, Actual Finish, % Complete, Resource, Units, Notes. Story Points (SP), the Constraint type and links
# 3-4 come from STORY_POINTS and EXTRA. (Cost is calculated on the Cost sheet from day rates.)
COLUMNS = ["Level", "Name", "Mode", "Dur", "P1", "T1", "L1", "P2", "T2", "L2", "Cons", "DL", "AS", "AF", "Pct", "Res",
           "Units", "Notes", "SP"]
_ = None
TASKS = [
    (1, "Initiation", "Auto", _, _, _, _, _, _, _, _, _, _, _, _, _, _, "Phase 1 - summary rows roll up dates, work, cost and progress automatically."),
    (2, "Project charter & scope statement", "Auto", 3, _, _, _, _, _, _, _, _, D(2026, 9, 7), D(2026, 9, 10), 1, "Priya Sharma", _, "Took 1 extra day for sponsor sign-off."),
    (2, "Stakeholder kickoff workshop", "Auto", 2, 2, "FS", _, _, _, _, _, _, D(2026, 9, 11), D(2026, 9, 14), 1, "Morgan Patel", _, ""),
    (2, "Kickoff complete", "Auto", 0, 3, "FS", _, _, _, _, _, _, D(2026, 9, 14), D(2026, 9, 14), 1, _, _, "Milestone (Duration 0)."),
    (1, "Requirements & Design", "Auto", _, _, _, _, _, _, _, _, _, _, _, _, _, _, ""),
    (2, "Requirements gathering", "Auto", 8, 4, "FS", _, _, _, _, _, _, D(2026, 9, 15), D(2026, 9, 25), 1, "Morgan Patel", _, ""),
    (2, "Solution architecture", "Auto", 5, 6, "SS", 3, _, _, _, _, _, D(2026, 9, 18), D(2026, 9, 25), 1, "Alex Chen", _, "Start-to-Start + 3 day lag example."),
    (2, "UX wireframes & prototype", "Auto", 6, 6, "FS", _, _, _, _, _, _, D(2026, 9, 28), D(2026, 10, 6), 1, "Casey Kim", _, ""),
    (2, "Design sign-off", "Auto", 0, 7, "FS", _, 8, "FS", _, _, _, D(2026, 10, 6), D(2026, 10, 6), 1, _, _, "Two predecessors - waits for the later one."),
    (1, "Build", "Auto", _, _, _, _, _, _, _, _, _, _, _, _, _, _, ""),
    (2, "Backend", "Auto", _, _, _, _, _, _, _, _, _, _, _, _, _, _, "Level-2 summary with sub-tasks."),
    (3, "Data model & database setup", "Auto", 5, 9, "FS", _, _, _, _, _, _, D(2026, 10, 7), D(2026, 10, 14), 1, "Jordan Lee", _, ""),
    (3, "Core API endpoints", "Auto", 10, 12, "FS", _, _, _, _, _, _, D(2026, 10, 15), _, 0.5, "Jordan Lee", _, "In progress - enter % Complete; forecast finish adjusts."),
    (3, "Payment gateway integration", "Auto", 6, 13, "FS", _, _, _, _, _, _, _, _, _, "Jordan Lee", _, ""),
    (2, "Frontend", "Auto", _, _, _, _, _, _, _, _, _, _, _, _, _, _, ""),
    (3, "UI component library", "Auto", 6, 9, "FS", _, _, _, _, _, _, D(2026, 10, 7), D(2026, 10, 16), 1, "Sam Rivera", _, ""),
    (3, "Feature screens", "Auto", 12, 16, "FS", _, 13, "SS", 2, _, _, D(2026, 10, 19), _, 0.25, "Sam Rivera", _, ""),
    (3, "Accessibility review", "Auto", 3, 17, "FF", 1, _, _, _, _, _, _, _, _, "Casey Kim", 0.5, "Finish-to-Finish + 1 day; 50% units."),
    (2, "Build complete", "Auto", 0, 14, "FS", _, 18, "FS", _, _, D(2026, 11, 6), _, _, _, _, _, "Milestone with a deadline of 06-Nov (see Deadline Check)."),
    (1, "Test & Release", "Auto", _, _, _, _, _, _, _, _, _, _, _, _, _, _, ""),
    (2, "System integration testing", "Auto", 8, 19, "FS", _, _, _, _, _, _, _, _, _, "Taylor Morgan", _, ""),
    (2, "User acceptance testing", "Auto", 6, 21, "FS", _, _, _, _, D(2026, 12, 7), _, _, _, _, "Morgan Patel", _, "Constraint: Start No Earlier Than 07-Dec (client availability)."),
    (2, "Training & user guides", "Auto", 5, 21, "SS", 2, _, _, _, _, _, _, _, _, "Casey Kim", 0.5, ""),
    (2, "Go-live readiness review", "Auto", 1, 22, "FS", _, 23, "FS", _, _, _, _, _, _, "Priya Sharma", _, "Three predecessors."),
    (2, "Go-live", "Auto", 0, 24, "FS", _, _, _, _, _, D(2026, 12, 18), _, _, _, _, _, "Milestone with a deadline of 18-Dec."),
    (2, "Hypercare support", "Auto", 10, 25, "FS", _, _, _, _, _, _, _, _, _, "Jordan Lee", 0.5, ""),
    (1, "Project Close", "Auto", _, _, _, _, _, _, _, _, _, _, _, _, _, _, ""),
    (2, "Steering committee update", "Manual", 1, _, _, _, _, _, _, D(2026, 11, 16), _, _, _, _, "Priya Sharma", _, "Manual mode example: fixed on its Constraint Date."),
    (2, "Lessons learned", "Auto", 2, 26, "FS", _, _, _, _, _, _, _, _, _, "Priya Sharma", _, ""),
    (2, "Project closed", "Auto", 0, 29, "FS", _, _, _, _, _, _, _, _, _, _, _, "Milestone"),
]


# Actual cost entered for a few tasks (by task number); the others use Cost x Progress.
# Story points (by task number) for the agile reports; milestones and summaries have none.
STORY_POINTS = {2: 3, 3: 2, 6: 8, 7: 5, 8: 5, 12: 5, 13: 8, 14: 5, 16: 5, 17: 8, 18: 3, 21: 8, 22: 5, 23: 3, 24: 1,
                26: 5, 28: 1, 29: 2}

# Who answers for each phase (by task number), shown on the Roadmap.
OWNERS = {1: "Priya Sharma", 5: "Morgan Patel", 10: "Jordan Lee", 20: "Taylor Morgan", 27: "Priya Sharma"}
# Other inputs (by task number): constraint types and a third predecessor.
EXTRA = {22: {"CType": "Start No Earlier Than"}, 24: {"P3": 21, "T3": "FS"}}
# More people on a task (beyond the first, in its Resources cell): (task number, resource, units). Task number = UID.
ASSIGNMENTS = [(7, "Jordan Lee", 0.5), (17, "Casey Kim", 0.5), (21, "Sam Rivera", 0.5), (22, "Taylor Morgan", None)]


# The plan as first estimated (Core API 8 days, Feature screens 10, before any actuals), frozen as the baseline of the
# committed sample (--baselined) so the Gantt shows planned against current on first open: UID -> (Baseline Start,
# Baseline Finish, Baseline Duration, Baseline Work), as the BL Source columns calculated them.
BASELINE = {
    1: (D(2026, 9, 7), D(2026, 9, 11), 5, 40),
    2: (D(2026, 9, 7), D(2026, 9, 9), 3, 24),
    3: (D(2026, 9, 10), D(2026, 9, 11), 2, 16),
    4: (D(2026, 9, 11), D(2026, 9, 11), 0, 0),
    5: (D(2026, 9, 14), D(2026, 10, 1), 14, 152),
    6: (D(2026, 9, 14), D(2026, 9, 23), 8, 64),
    7: (D(2026, 9, 17), D(2026, 9, 23), 5, 40),
    8: (D(2026, 9, 24), D(2026, 10, 1), 6, 48),
    9: (D(2026, 10, 1), D(2026, 10, 1), 0, 0),
    10: (D(2026, 10, 2), D(2026, 10, 28), 19, 292),
    11: (D(2026, 10, 2), D(2026, 10, 28), 19, 152),
    12: (D(2026, 10, 2), D(2026, 10, 8), 5, 40),
    13: (D(2026, 10, 9), D(2026, 10, 20), 8, 64),
    14: (D(2026, 10, 21), D(2026, 10, 28), 6, 48),
    15: (D(2026, 10, 2), D(2026, 10, 27), 18, 140),
    16: (D(2026, 10, 2), D(2026, 10, 9), 6, 48),
    17: (D(2026, 10, 13), D(2026, 10, 26), 10, 80),
    18: (D(2026, 10, 23), D(2026, 10, 27), 3, 12),
    19: (D(2026, 10, 28), D(2026, 10, 28), 0, 0),
    20: (D(2026, 10, 29), D(2026, 12, 29), 44, 180),
    21: (D(2026, 10, 29), D(2026, 11, 9), 8, 64),
    22: (D(2026, 12, 7), D(2026, 12, 14), 6, 48),
    23: (D(2026, 11, 2), D(2026, 11, 6), 5, 20),
    24: (D(2026, 12, 15), D(2026, 12, 15), 1, 8),
    25: (D(2026, 12, 15), D(2026, 12, 15), 0, 0),
    26: (D(2026, 12, 16), D(2026, 12, 29), 10, 40),
    27: (D(2026, 11, 16), D(2026, 12, 31), 34, 24),
    28: (D(2026, 11, 16), D(2026, 11, 16), 1, 8),
    29: (D(2026, 12, 30), D(2026, 12, 31), 2, 16),
    30: (D(2026, 12, 31), D(2026, 12, 31), 0, 0),
}


def assignments(blank=False, stress=0):
    """Assignment rows for the sample; stress=N repeats them for each copy (UIDs follow the copies)."""
    if blank:
        return []
    n = len(TASKS)
    return [(uid + copy * n, res, units) for copy in range(max(1, stress)) for uid, res, units in ASSIGNMENTS]


def _row(i, t):
    return dict(zip(COLUMNS, t), SP=STORY_POINTS.get(i + 1), Owner=OWNERS.get(i + 1), **EXTRA.get(i + 1, {}))


def tasks(blank=False, stress=0, baselined=False):
    """Rows as dicts keyed by COLUMNS. stress=N repeats the sample N times; each copy follows the previous one.
    baselined: the rows carry BASELINE (the committed sample), so planned and current differ."""
    if blank:
        return []
    if not stress:
        rows = [take_links(dict(_row(i, t), UID=i + 1)) for i, t in enumerate(TASKS)]
        if baselined:
            for r in rows:
                r.update(zip(("BS", "BF", "BD", "BW"), BASELINE[r["UID"]]))
        return rows
    rows = []
    n = len(TASKS)
    for copy in range(stress):
        for i, t in enumerate(TASKS):
            row = _row(i, t)
            for p in ("P1", "P2", "P3", "P4"):
                if row.get(p):
                    row[p] += copy * n
            if copy and row["Level"] == 2 and row["Name"] == "Project charter & scope statement":
                row["P1"], row["T1"] = copy * n, "FS"      # chain the copies: each starts after the previous ends
            row["Name"] = f"{row['Name']} ({copy + 1})"
            row["UID"] = copy * n + i + 1
            rows.append(take_links(row))
    return rows


# Release 1.0: the sample types two durations in weeks (the 3 / 3d / 2w input) and picks every link and person
RELEASE_DURATIONS = {23: "1w", 26: "2w"}


def to_pickers(row):
    """A sample row with its Predecessors / Resources text moved into the Pred1-4 / Type / Lag and Res1-4 / Units
    pickers (Release 1.0), so the sample shows the pickers; the typed text cells stay blank."""
    row = dict(row)
    if row.get("UID") in RELEASE_DURATIONS:
        row["Dur"] = RELEASE_DURATIONS[row["UID"]]
    links = [p.strip() for p in re.split(r"[,;]", row.pop("PredIn", None) or "") if p.strip()]
    for k, text in enumerate(links[:LINKS], start=1):
        m = re.match(r"^(\d+)(FS|SS|FF|SF)?([+-]\d+)?d?$", text, re.I)
        row[f"Pick{k}"] = int(m.group(1))
        if m.group(2) and m.group(2).upper() != "FS":
            row[f"PickT{k}"] = m.group(2).upper()
        if m.group(3):
            row[f"PickL{k}"] = int(m.group(3))
    people = [p.strip() for p in (row.pop("ResIn", None) or "").split(",") if p.strip()]
    for k, text in enumerate(people[:PEOPLE], start=1):
        who, _, units = text.partition("[")
        row[f"Who{k}"] = who.strip()
        if units:
            row[f"WhoU{k}"] = float(units.rstrip("]%")) / 100
    return row


# The Release 1.0 Simple sample: 12 tasks of the full one (UIDs kept), linked among themselves (a value replaces a
# link to a task left out)
SIMPLE = {1: None, 2: None, 3: None, 4: None, 5: None, 6: None, 8: "6", 9: "8", 20: None, 21: "9", 24: "21", 25: "24"}
# its baseline, as its BL Source columns calculate it before any actuals
SIMPLE_BASELINE = {
    1: (D(2026, 9, 7), D(2026, 9, 11), 5, 40), 2: (D(2026, 9, 7), D(2026, 9, 9), 3, 24),
    3: (D(2026, 9, 10), D(2026, 9, 11), 2, 16), 4: (D(2026, 9, 11), D(2026, 9, 11), 0, 0),
    5: (D(2026, 9, 14), D(2026, 10, 1), 14, 112), 6: (D(2026, 9, 14), D(2026, 9, 23), 8, 64),
    8: (D(2026, 9, 24), D(2026, 10, 1), 6, 48), 9: (D(2026, 10, 1), D(2026, 10, 1), 0, 0),
    20: (D(2026, 10, 2), D(2026, 10, 23), 16, 128), 21: (D(2026, 10, 2), D(2026, 10, 22), 15, 120),
    24: (D(2026, 10, 23), D(2026, 10, 23), 1, 8), 25: (D(2026, 10, 23), D(2026, 10, 23), 0, 0),
}


def simple_tasks(baselined=False):
    """The 12-task sample of the Release 1.0 Simple workbook."""
    rows = []
    for r in tasks(baselined=False):
        if r["UID"] not in SIMPLE:
            continue
        r = dict(r)
        if SIMPLE[r["UID"]]:                                  # a link to a task the short plan leaves out
            r["PredIn"] = SIMPLE[r["UID"]]
        r.pop("SP", None)
        if r["UID"] == 21:                                   # in the short plan, testing follows design sign-off
            r.update(Dur=15, AS=D(2026, 10, 7), Stage="In Testing", Notes="Started a few days late; on track.")
        if baselined:
            r.update(zip(("BS", "BF", "BD", "BW"), SIMPLE_BASELINE[r["UID"]]))
        rows.append(r)
    return rows


# template 1.1: RACI (Accountable, Consulted, Informed), three-point estimates (Optimistic, Pessimistic) and one
# what-if Sandbox Duration (a second developer on the core API), by UID. Settings > Sandbox starts at No, so the plan
# schedules as in Release 1.0.
V11_TASKS = {
    2: {"Acc": "Priya Sharma", "Con": "Morgan Patel", "Inf": "Steering committee"},
    13: {"Acc": "Jordan Lee", "Con": "Alex Chen", "Inf": "Priya Sharma", "Opt": 8, "Pess": 16, "SandDur": 7},
    17: {"Acc": "Sam Rivera", "Con": "Casey Kim", "Inf": "Priya Sharma", "Opt": 10, "Pess": 18},
    21: {"Acc": "Taylor Morgan", "Con": "Jordan Lee, Sam Rivera", "Inf": "Morgan Patel", "Opt": 6, "Pess": 13},
    22: {"Opt": 5, "Pess": 10},
    25: {"Acc": "Priya Sharma", "Con": "Steering committee", "Inf": "All staff"},
}


def with_v11(row):
    """A sample row with its template 1.1 inputs."""
    return {**row, **V11_TASKS.get(row.get("UID"), {})}


# template 1.2: custom fields (B-509) on a few tasks: a workstream, a risk score and a "needs sign-off" flag
V12_TASKS = {13: {"CText1": "Payments", "CNum1": 4, "CFlag1": "Yes"}, 14: {"CText1": "Payments", "CNum1": 5},
             17: {"CText1": "Web", "CNum1": 2}, 22: {"CText1": "Client", "CFlag1": "Yes"}}


def with_v12(row):
    """A sample row with its template 1.2 inputs."""
    return {**row, **V12_TASKS.get(row.get("UID"), {})}
