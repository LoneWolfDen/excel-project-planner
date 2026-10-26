"""Instructions sheet."""
from openpyxl.styles import Alignment

from .core import CALC_FILL, HEADER_FILLS, HEADER_FONT, INPUT_FONT, font


def text(n_tasks):
    return [
        ("HOW IT WORKS", None),
        ("1. Settings", "Set project start, working week (Mon-Sun Yes/No), hours/day, status date, Gantt scale and sprint "
                        "calendar. Add holidays on 'Holidays'."),
        ("2. Resources", "List people/roles with Max Units and hourly rate. They appear in the Resource dropdown on Tasks and "
                         "in sprint capacity."),
        ("3. Tasks", "One row per task. Enter Level (1 = phase/summary, 2 = task, 3 = sub-task...), Task Name, Duration in "
                     "working days (0 = milestone), and up to two predecessors by UID with link type (FS, SS, FF, SF) and lag in "
                     "days (negative = lead). A row becomes a SUMMARY automatically when the rows beneath it have a higher "
                     "level; its dates, work, cost and progress roll up. Optional Deadline: Deadline Check shows 'Late by Nd'. "
                     "Hover over a column header for what it means."),
        ("4. Set the baseline", "When the plan is agreed: on Tasks, select the 5 'BL Source' columns (all task rows), Copy, then "
                                "select the first 'Baseline Start' cell and Paste Special > Values. Planned Start/Finish now stay "
                                "frozen (like MS Project's Baseline) and the variance columns compare against it. To re-baseline, "
                                "repeat; to clear it, delete the 5 Baseline columns' contents."),
        ("5. Track progress", "Enter Actual Start, Actual Finish and % Complete. Forecast Start/Finish then use actuals and push "
                              "every dependent task. Start / Finish / Duration / Work / Cost Var and Status flag slippage."),
        ("6. Gantt", "Read-only view: dark outline = planned dates, green = done, blue = remaining on track, orange = remaining "
                     "but late vs plan, dark grey = summary, purple = milestone, red cell = deadline, red border = status date. "
                     "Change GanttStart / GanttPeriod (1 = days, 7 = weeks) on Settings to scroll or zoom."),
        ("7. Sprints", "Sprint calendar from Settings. Each task is tagged with the sprint its forecast start falls in. Work "
                       "hours are prorated across sprints per resource, compared against capacity (Max Units x Hours/day x "
                       "working days); over-allocations turn red."),
        ("8. Dashboard", "KPIs (including baseline, missed deadlines, link problems), status counts, upcoming milestones and "
                         "sprint load vs capacity."),
        ("CELL COLOURS", None),
        ("Blue text", "Input - type here."),
        ("Black text on grey", "Formula - do not overwrite. Summary rows grey-out inputs that are ignored."),
        ("Dark blue header", "Input column.  Dark grey header = calculated.  Light grey header = helper column (hidden, grouped)."),
        ("RULES (same as MS Project)", None),
        ("ID and UID", "ID = position of the task (renumbers when rows move, like MS Project). UID = permanent number that "
                       "predecessors refer to, so links survive inserting, deleting and moving rows. A duplicate or missing UID "
                       "turns red."),
        ("Inserting tasks", "Tasks is an Excel Table: right-click a row > Insert > Table Rows Above (or type in the row under "
                            "the table). Every formula fills in automatically. Set the new task's UID to Settings > NextFreeUID "
                            "(also shown at the top of Tasks)."),
        ("Outline levels", "Each row's Level may be at most 1 higher than the row above. A '?' in WBS (red) means a level jump."),
        ("Predecessors", "Enter the predecessor's UID. Links to the task itself, to its own summary, to an empty row or to a "
                         "missing UID are ignored and explained in the Link Check column (red). Summary rows ignore Duration / "
                         "Predecessors / Actuals - they roll up. A loop through several tasks (A > B > A) still causes a "
                         "circular reference."),
        ("Mode", "Auto = scheduled from predecessors, project start and 'Start No Earlier Than'. Manual = pinned to the date in "
                 "'Start No Earlier Than'."),
        ("Rescheduling", "Settings > RescheduleAfterStatus = Yes moves incomplete work to after the status date (like MS "
                         "Project 'Update Project'), so late tasks push successors."),
        ("Capacity", f"Tables grow when you add rows: Tasks starts with {n_tasks} rows, Resources 50, Sprints 26, Holidays 100. "
                     f"The Gantt and MSP_Tasks sheets show the first {n_tasks} tasks (the Gantt says so when there are more)."),
        ("Speed", "A typical edit recalculates in a fraction of a second. Each change to the build is timed in CI; see "
                  "docs/ARCHITECTURE.md in the repository. For large plans: enter a Status date instead of leaving it on "
                  "TODAY, and keep the Gantt to the period you need."),
        ("MOVING TO MICROSOFT PROJECT", None),
        ("One-time setting", "MS Project > File > Options > Trust Center > Trust Center Settings > Legacy Formats > 'Allow "
                             "loading files with legacy or non-default file formats'."),
        ("Import", "In MS Project: File > Open > choose this .xlsx > Import Wizard > New map > 'As a new project' > tick Tasks "
                   "(and Resources) > Tasks source worksheet = MSP_Tasks, Resources = MSP_Resources. Column headers already "
                   "use MS Project field names, so the wizard maps them automatically. Map 'Predecessors' (by ID) for a new "
                   "project; 'Unique ID' and 'Unique ID Predecessors' are for merging into an existing one."),
        ("Tip: Start/Finish", "For auto-scheduled tasks map Duration + Predecessors (+ Constraint Type/Date) and set Start/Finish "
                              "to '(not mapped)' - importing Start/Finish makes MS Project add fixed 'Start No Earlier Than' "
                              "constraints. Map Baseline Start/Finish to keep the plan, and Actual Start/Finish/% Complete for "
                              "progress."),
        ("Rows", "Blank MSP_Tasks rows import as blank tasks - delete them in MS Project, or filter the table to non-blank "
                 "names and copy those rows to a new sheet before importing."),
        ("From MS Project to Excel", "MS Project > File > Save As > Excel Workbook, then paste Name / Outline Level / Duration "
                                     "/ Predecessors etc. into the Tasks input columns (split '3FS+2d' into UID, Type, Lag)."),
        ("NAMES AND TABLES (Formulas > Name Manager)", None),
        ("Tables", "Tasks, ResourceTable, SprintTable, HolidayTable, GanttView, MSPTasks, MSPResources, KPIs, StatusSummary, "
                   "NextMilestones and the List* dropdown tables. Inside a table, formulas read [@Column] (this row)."),
        ("Settings", "ProjectStart, StatusDate, StatusWorkday, WorkWeek, Holidays, HoursPerDay, DefaultUnits, "
                     "RescheduleAfterStatus, GanttStart, GanttPeriod, Sprint1Start, SprintLength, SprintCount, SprintPrefix, "
                     "SprintLastEnd, NextFreeUID"),
        ("Task columns", "tsk_ID, tsk_UID, tsk_WBS, tsk_Name, tsk_Duration, tsk_PlanStart, tsk_FcFinish, tsk_Status, tsk_Resource, "
                         "tsk_Work, tsk_Cost ... each is Tasks[Column], so it grows with the table. Other sheets read Tasks "
                         "only through these names."),
        ("Resources / Sprints", "res_Name, res_MaxUnits, res_Rate, ...  spr_Name, spr_Start, spr_End, spr_WorkDays, spr_Capacity, "
                                "spr_Work"),
        ("Key formulas", "Calculated start = WORKDAY.INTL(MAX(ProjectStart, [@Start No Earlier Than], predecessor drivers) - 1, "
                         "1, WorkWeek, Holidays).  Predecessor driver (FS) = WORKDAY.INTL(OFFSET(Tasks[[#Headers],[_CalcFinish]], "
                         "MATCH([@Pred1 UID], Tasks[UID], 0), 0), 1 + Lag, WorkWeek, Holidays).  Planned Start = Baseline Start "
                         "if set, else calculated.  Summary start = MIN of the rows below it, OFFSET(header, [@_Row]+1, 0, "
                         "[@_Descendants], 1), where _Descendants = COUNTIF(Tasks[WBS], [@WBS] & \".*\")."),
    ]


def build_instructions(wb, n_tasks):
    ws = wb.create_sheet("Instructions")
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 110
    ws["B1"] = "Project Planner - MS Project-compatible Excel template"
    ws["B1"].font = font(size=18, bold=True, color="1F4E78")
    r = 3
    for label, body in text(n_tasks):
        if body is None:
            r += 1
            ws.cell(r, 2, label).font = HEADER_FONT
            ws.cell(r, 2).fill = HEADER_FILLS["in"]
            ws.cell(r, 3).fill = HEADER_FILLS["in"]
        else:
            ws.cell(r, 2, label).font = font(bold=True)
            ws.cell(r, 2).alignment = Alignment(vertical="top")
            c = ws.cell(r, 3, body)
            c.font = INPUT_FONT if label == "Blue text" else font()
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if label == "Black text on grey":
                c.fill = CALC_FILL
        r += 1
    return ws
