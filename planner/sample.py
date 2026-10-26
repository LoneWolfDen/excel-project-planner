"""Sample project, and a generator that repeats it to make large stress-test workbooks."""
import datetime as dt

D = dt.date

# Level, Task Name, Mode, Duration, Pred1 UID, Type, Lag, Pred2 UID, Type, Lag, Start No Earlier Than, Deadline,
# Actual Start, Actual Finish, % Complete, Resource, Units, Notes. Actual Cost (ACin) and Story Points (SP) come
# from ACTUAL_COST and STORY_POINTS.
COLUMNS = ["Level", "Name", "Mode", "Dur", "P1", "T1", "L1", "P2", "T2", "L2", "Cons", "DL", "AS", "AF", "Pct", "Res",
           "Units", "Notes", "ACin", "SP"]
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
    (2, "User acceptance testing", "Auto", 6, 21, "FS", _, _, _, _, D(2026, 12, 7), _, _, _, _, "Morgan Patel", _, "Start No Earlier Than 07-Dec (client availability)."),
    (2, "Training & user guides", "Auto", 5, 21, "SS", 2, _, _, _, _, _, _, _, _, "Casey Kim", 0.5, ""),
    (2, "Go-live readiness review", "Auto", 1, 22, "FS", _, 23, "FS", _, _, _, _, _, _, "Priya Sharma", _, ""),
    (2, "Go-live", "Auto", 0, 24, "FS", _, _, _, _, _, D(2026, 12, 18), _, _, _, _, _, "Milestone with a deadline of 18-Dec."),
    (2, "Hypercare support", "Auto", 10, 25, "FS", _, _, _, _, _, _, _, _, _, "Jordan Lee", 0.5, ""),
    (1, "Project Close", "Auto", _, _, _, _, _, _, _, _, _, _, _, _, _, _, ""),
    (2, "Steering committee update", "Manual", 1, _, _, _, _, _, _, D(2026, 11, 16), _, _, _, _, "Priya Sharma", _, "Manual mode example: fixed on the date in the Start No Earlier Than column."),
    (2, "Lessons learned", "Auto", 2, 26, "FS", _, _, _, _, _, _, _, _, _, "Priya Sharma", _, ""),
    (2, "Project closed", "Auto", 0, 29, "FS", _, _, _, _, _, _, _, _, _, _, _, "Milestone"),
]


# Actual cost entered for a few tasks (by task number); the others use Cost x Progress.
ACTUAL_COST = {2: 3040, 6: 5040, 13: 4080}
# Story points (by task number) for the agile reports; milestones and summaries have none.
STORY_POINTS = {2: 3, 3: 2, 6: 8, 7: 5, 8: 5, 12: 5, 13: 8, 14: 5, 16: 5, 17: 8, 18: 3, 21: 8, 22: 5, 23: 3, 24: 1,
                26: 5, 28: 1, 29: 2}


def _row(i, t):
    return dict(zip(COLUMNS, t), ACin=ACTUAL_COST.get(i + 1), SP=STORY_POINTS.get(i + 1))


def tasks(blank=False, stress=0):
    """Rows as dicts keyed by COLUMNS. stress=N repeats the sample N times; each copy follows the previous one."""
    if blank:
        return []
    if not stress:
        return [_row(i, t) for i, t in enumerate(TASKS)]
    rows = []
    n = len(TASKS)
    for copy in range(stress):
        for i, t in enumerate(TASKS):
            row = _row(i, t)
            for p in ("P1", "P2"):
                if row[p]:
                    row[p] += copy * n
            if copy and row["Level"] == 2 and row["Name"] == "Project charter & scope statement":
                row["P1"], row["T1"] = copy * n, "FS"      # chain the copies: each starts after the previous ends
            row["Name"] = f"{row['Name']} ({copy + 1})"
            rows.append(row)
    return rows
