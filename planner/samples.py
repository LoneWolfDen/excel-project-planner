"""More sample projects (template 1.2, B-804): `--sample construction` and `--sample event` build the planner around
a small realistic plan of another kind, as an import would (planner.build.ProjectData), so they use every sheet
and show the project type's stages. The software sample stays planner/sample.py.

Tasks are written as an outline would type them: Level, name, duration, predecessors by UID ("2", "5SS+3d"),
people, and progress for the work already done at the sample's status date (25-Oct-2026).
"""
import datetime as dt

D = dt.date
# (level, name, duration, predecessors, resources, stage, actual start, actual finish, deadline)
CONSTRUCTION = {
    "settings": {"ProjectName": "Riverside Clinic Extension", "ProjectManager": "Dana Okafor",
                 "ProjectStartDate": D(2026, 9, 7), "StatusDateInput": D(2026, 10, 25),
                 "ProjectType": "Construction"},
    "people": [("Dana Okafor", "DO", "Project Manager", 0.5, 820), ("Luis Ortega", "LO", "Site Manager", 1, 640),
               ("Grace Mensah", "GM", "Structural Engineer", 0.4, 900), ("Ravi Kumar", "RK", "Electrician", 1, 520),
               ("Ana Silva", "AS", "Plumber", 1, 500), ("Tom Becker", "TB", "Carpenter", 1, 480)],
    "tasks": [
        (1, "Pre-construction", None, None, None, None, None, None, None),
        (2, "Planning permission", 10, None, "Dana Okafor", "Complete", D(2026, 9, 7), D(2026, 9, 18), None),
        (2, "Site survey", 3, None, "Luis Ortega, Grace Mensah", "Complete", D(2026, 9, 7), D(2026, 9, 9), None),
        (2, "Contractor appointed", 0, "2, 3", None, "Complete", D(2026, 9, 18), D(2026, 9, 18), None),
        (1, "Groundworks", None, None, None, None, None, None, None),
        (2, "Site set-up and fencing", 3, "4", "Luis Ortega", "Complete", D(2026, 9, 21), D(2026, 9, 23), None),
        (2, "Excavation", 8, "6", "Luis Ortega", "Complete", D(2026, 9, 24), D(2026, 10, 6), None),
        (2, "Foundations poured", 6, "7", "Luis Ortega, Grace Mensah[50%]", "Inspection", D(2026, 10, 7), None,
         None),
        (2, "Foundations signed off", 0, "8", None, None, None, None, D(2026, 10, 23)),
        (1, "Structure", None, None, None, None, None, None, None),
        (2, "Steel frame", 10, "9", "Luis Ortega, Grace Mensah[40%]", None, None, None, None),
        (2, "Roof", 8, "11", "Tom Becker", None, None, None, None),
        (2, "Building watertight", 0, "12", None, None, None, None, D(2026, 12, 11)),
        (1, "Fit-out", None, None, None, None, None, None, None),
        (2, "Electrical first fix", 6, "13", "Ravi Kumar", None, None, None, None),
        (2, "Plumbing first fix", 6, "13", "Ana Silva", None, None, None, None),
        (2, "Partitions and doors", 8, "15, 16", "Tom Becker", None, None, None, None),
        (2, "Second fix and finishes", 10, "17", "Ravi Kumar, Ana Silva, Tom Becker", None, None, None, None),
        (1, "Handover", None, None, None, None, None, None, None),
        (2, "Snagging", 5, "18", "Luis Ortega", None, None, None, None),
        (2, "Building control completion", 2, "20", "Dana Okafor", None, None, None, None),
        (2, "Handover to the clinic", 0, "21", "Dana Okafor", None, None, None, D(2027, 2, 26)),
    ],
}
EVENT = {
    "settings": {"ProjectName": "Northwind Partner Summit 2027", "ProjectManager": "Mia Laurent",
                 "ProjectStartDate": D(2026, 9, 7), "StatusDateInput": D(2026, 10, 25),
                 "ProjectType": "Software / Agile"},
    "people": [("Mia Laurent", "ML", "Event Lead", 1, 700), ("Owen Price", "OP", "Logistics", 1, 450),
               ("Priya Nair", "PN", "Marketing", 1, 520), ("Ken Ito", "KI", "Speaker Liaison", 1, 480),
               ("Zoe Adams", "ZA", "Designer", 1, 550)],
    "tasks": [
        (1, "Concept", None, None, None, None, None, None, None),
        (2, "Goals, audience and budget", 5, None, "Mia Laurent", "Complete", D(2026, 9, 7), D(2026, 9, 11), None),
        (2, "Venue shortlist and visits", 10, "2", "Owen Price", "Complete", D(2026, 9, 14), D(2026, 9, 25), None),
        (2, "Venue contract signed", 0, "3", None, "Complete", D(2026, 9, 28), D(2026, 9, 28), None),
        (1, "Programme", None, None, None, None, None, None, None),
        (2, "Call for speakers", 15, "4", "Ken Ito", "In Testing", D(2026, 9, 29), None, None),
        (2, "Agenda published", 0, "6", None, None, None, None, D(2026, 11, 13)),
        (2, "Speaker briefings", 5, "7", "Ken Ito", None, None, None, None),
        (1, "Marketing", None, None, None, None, None, None, None),
        (2, "Brand and website", 10, "4", "Zoe Adams, Priya Nair[50%]", "In Development", D(2026, 10, 1), None,
         None),
        (2, "Registration opens", 0, "10", None, None, None, None, D(2026, 11, 2)),
        (2, "Email campaign", 20, "11", "Priya Nair", None, None, None, None),
        (1, "Logistics", None, None, None, None, None, None, None),
        (2, "Catering and AV booked", 5, "4", "Owen Price", "Complete", D(2026, 10, 1), D(2026, 10, 9), None),
        (2, "Travel and hotels for speakers", 8, "7", "Owen Price", None, None, None, None),
        (2, "Signage and badges printed", 4, "8, 12SS+10d", "Zoe Adams, Owen Price", None, None, None, None),
        (1, "Event", None, None, None, None, None, None, None),
        (2, "Rehearsal day", 1, "16", "Mia Laurent, Ken Ito", None, None, None, None),
        (2, "Summit (two days)", 2, "18", "Mia Laurent, Owen Price, Ken Ito, Priya Nair", None, None, None,
         D(2027, 1, 29)),
        (2, "Survey and wrap-up report", 5, "19", "Mia Laurent", None, None, None, None),
    ],
}
# other inputs by UID: the summit is booked for fixed dates, the rehearsal the day before (a date alone = Start
# No Earlier Than)
EVENT["extra"] = {18: {"Cons": D(2027, 1, 26)}, 19: {"CType": "Must Start On", "Cons": D(2027, 1, 27)}}
SAMPLES = {"construction": CONSTRUCTION, "event": EVENT}


def project(which):
    """The ProjectData for a named sample."""
    from .build import ProjectData
    s = SAMPLES[which]
    rows = []
    for uid, (level, name, dur, pred, res, stage, act_start, act_finish, deadline) in enumerate(s["tasks"], start=1):
        row = {"UID": uid, "Level": level, "Name": name}
        for key, value in (("Dur", dur), ("PredIn", pred), ("ResIn", res), ("Stage", stage), ("AS", act_start),
                           ("AF", act_finish), ("DL", deadline)):
            if value is not None:
                row[key] = value
        row.update(s.get("extra", {}).get(uid, {}))
        rows.append(row)
    people = [p + (f"{p[1].lower()}@example.com",) for p in s["people"]]
    return ProjectData(settings=dict(s["settings"]), tasks=rows, resources=people)
