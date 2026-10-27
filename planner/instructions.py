"""Start sheet: the six settings a plan needs (written by planner.settings), a quick start, the weekly update, and
the full reference below them."""
from openpyxl.styles import Alignment
from openpyxl.worksheet.hyperlink import Hyperlink

from .core import CALC_FILL, HEADER_FILLS, HEADER_FONT, INPUT_FONT, NOTE_FONT, font
from .settings import ESSENTIAL, START, START_ROW

# (label, text, sheet to link to)
QUICK_START = [
    ("1. Your project", "Fill in the blue cells above: name, manager, start date, working week and hours per day. "
                        "Everything else has a sensible default on Settings.", None),
    ("2. People", "List the people (or roles) with their availability, hourly rate and any leave (by month, on the "
                  "same row).", "Resources"),
    ("3. Tasks", "One row per task: Level (1 = phase, 2 = task ...), Task Name, Duration in working days, the UID of "
                 "the tasks that must come first (Predecessors, e.g. 12 or 12, 14SS+2d) and who does it (Resources). That's a plan: dates, the "
                 "critical path, the Gantt and the reports follow. Optional columns are grouped: click + above "
                 "a section to show them.", "Tasks"),
    ("4. Check", "Fix anything in the Issues column (red), then look at the Gantt and the Roadmap.", "Gantt"),
    ("5. Baseline", "When the plan is agreed, freeze it: on Tasks open the BASELINE section (+), copy the 5 'BL "
                    "Source' columns and Paste Special > Values onto the 5 Baseline columns.", "Tasks"),
]
WEEKLY = [
    ("Status sheet", "The weekly update starts on the Status sheet: type the status date at its top, then follow "
                     "the steps under it (progress on Tasks, the Log, the RAG, the commentary, the History "
                     "snapshot).", "Status"),
]


def text(sizes, wizard=False):
    return [
        ("REQUIREMENTS", None),
        ("Excel version", "Excel 2019 or Microsoft 365 (Windows or Mac), or Excel for the web. No macros needed."),
        ("HOW IT WORKS", None),
        ("1. Settings", "Project start (a typed date; Ctrl+; enters today), working week and hours/day are at the "
                        "top of this sheet; the status date is at the top of the Status sheet. Settings holds the rest (Gantt scale, sprint calendar, "
                        "RAG thresholds ...). Holidays are listed on Settings, below the settings. Update the status date whenever you report "
                        "progress."),
        ("2. Resources", "List people/roles with Max Units and hourly rate. Type their names (or initials) in the "
                         "Resources column on Tasks: several people are comma-separated, with an optional share of "
                         "their time, e.g. 'Alex Chen, JL[50%]'."),
        ("3. Tasks", "One row per task. Enter Level (1 = phase/summary, 2 = task, 3 = sub-task...), Task Name, Duration in "
                     "working days (0 = milestone), and Predecessors: the UIDs of up to four tasks that must come first, comma-separated, "
                     "each with an optional link type (FS, SS, FF, SF) and lag in days (negative = lead), e.g. "
                     "'12, 14SS+2d'. A row becomes a SUMMARY automatically when the rows beneath it have a higher "
                     "level; its dates, work, cost and progress roll up. Optional Deadline: Deadline Check shows 'Late by Nd'. "
                     "Hover over a column header for what it means."),
        ("4. Set the baseline", "When the plan is agreed: on Tasks, select the 5 'BL Source' columns (all task rows), Copy, then "
                                "select the first 'Baseline Start' cell and Paste Special > Values. Planned Start/Finish now stay "
                                "frozen (like MS Project's Baseline) and the variance columns compare against it. To re-baseline, "
                                "repeat; to clear it, delete the 5 Baseline columns' contents."),
        ("5. Track progress", "Pick each task's Stage (Not Started, Design, In Development, Test Ready, In Testing, "
                              "Complete; each stage's % is on Settings > Stages, editable) and enter Actual Start once "
                              "it starts and Actual Finish when it's done; Issues asks for the dates the Stage calls "
                              "for. % Complete (DETAIL section) overrides the stage's % when typed. Start / Finish (the current schedule) then use actuals and push "
                              "every dependent task. Start / Finish / Duration / Work / Cost Var and Status flag slippage. "
                              "Optionally enter Actual Cost (cost to date) on detail tasks; blank = Cost x Progress."),
        ("Critical path", "Late Start / Late Finish come from a backward pass from the forecast project finish (and any "
                          "deadline). Total Slack = how far a task can slip before it delays the project or a deadline; Free "
                          "Slack = before it delays a successor. Critical = slack <= Settings > CriticalSlackDays on an "
                          "unfinished task. A missed deadline gives negative slack to the task and the chain feeding it."),
        ("6. Gantt", "Read-only view. The panel shows ID, Task, Start, Finish, % and Resources; click + above it for WBS, "
                     "duration, planned dates and status. Timeline: black outline = planned dates (the baseline once set; a "
                     "task added after the baseline has no outline), pale blue-grey = planned days no work is on now, "
                     "green = done, blue = remaining on track, red = remaining past the planned finish (slip), orange = "
                     "remaining on the critical path, dark grey = summary, purple = milestone, red cell = deadline, red "
                     "line = status date. "
                     "Change GanttStart / GanttPeriod (1 = days, 7 = weeks, 30 = months, 91 = quarters) on Settings to "
                     "scroll or zoom. Collapse the outline with the 1 2 3 buttons at the top left (rows are grouped by "
                     "Level), or set Settings > OutlineShowLevel and filter Show = Yes (also on Tasks)."),
        ("6b. Roadmap", "One landscape page for sponsors: phases and milestones with their dates, drawn by month (blue "
                        "= forecast, green = done, amber = late vs plan, outline = planned, purple = milestone, red = "
                        "deadline missed). Settings > RoadmapStart scrolls the 18 months shown."),
        ("7. Agile", "Sprint calendar from Settings. Each task is tagged with the sprint its forecast start falls in. Work "
                     "hours are prorated across sprints and compared against capacity (Max Units x Hours/day x "
                     "working days, less leave). Optional agile columns: Story Points and Planned Sprint on Tasks "
                     "(Planned Sprint wins over the automatic sprint), Points Capacity per sprint; the sheet then shows "
                     "committed and done points, velocity, the release forecast, the current sprint's backlog and "
                     "daily burndown, and the charts. Each person's load by sprint is on the Resource Report."),
        ("8. Status", "The weekly page, printable on one page: the status date and this week's steps, overall RAG and "
                      "RAG by area (schedule, SPI, CPI, deadlines, critical path, RAID, plan quality) calculated from "
                      "the thresholds on Settings (RAG...); pick a PM RAG to override one and say why in Comment. Then "
                      "key figures, tasks by status, next milestones, earned value (BAC, PV, EV, AC, SPI, CPI, EAC, "
                      "VAC, TCPI), the top open RAID items, change control, the commentary for the period and the "
                      "snapshot row for History."),
        ("9. Schedule Report", "Phase summary with RAG, all milestones, late tasks, tasks starting and due in the "
                                "look-ahead window (Settings > LookaheadDays) and open critical tasks. Lists update with "
                                "every edit; drag a list's bottom corner down to show more rows."),
        ("10. Resource Report", "Per resource: work, done, remaining, cost, actual cost, late and critical tasks, peak "
                                "week and weeks over capacity. Hours and utilisation grids by week from "
                                "Settings > ResourceWeekStart and by sprint (red = over capacity), leave by week, and a "
                                "team hours vs capacity chart."),
        ("Leave", "On Resources, to the right of each person's details: one column per month (Jan-26 ... Dec-28; add a month by inserting a column "
                  "among the months and typing e.g. Jan-29). Type the days off: '16, 17(0.5)' under Nov-26 = all of 16 November "
                  "and half of the 17th. Weekends and holidays are ignored. Leave lowers capacity on the Resource "
                  "Report and the Agile sheet and can make a task Over-allocated; it does not move task dates."),
        ("11. Log", "One list for risks, assumptions, issues, dependencies (RAID), decisions and change requests: "
                    "pick the Type, then fill in the columns that aren't grey. RAID: probability and impact (1-5); "
                    "Score = probability x impact (issues count as probability 5); Rating from Settings > "
                    "RAIDHighScore / RAIDMediumScore. Decisions and changes: options and reasons, who decided and "
                    "when; a change's schedule and cost impact. Overdue when still open past its due date. Link a "
                    "task by UID and another item by its Ref. The Status sheet lists the top open RAID items, rates "
                    "RAID amber when any open one is High and totals approved and pending changes. Approving a "
                    "change doesn't edit the plan: update the tasks, and re-baseline if the sponsor agrees."),
        ("12. History", "On each reporting date, copy the SnapshotNow row at the bottom of the Status sheet and Paste "
                        "Special > Values onto the first empty row of History's StatusHistory. Trend charts: forecast vs planned finish, earned value "
                        "S-curve, SPI / CPI and a story-point burn-up that shows scope changes."),
        ("Earned value", "Set a baseline first: without one, Budget = current Cost and PV follows the current plan, so "
                         "SPI and CPI only show progress against the latest plan. PV = budget of work planned before the "
                         "status date; EV = Budget x Progress; AC = Actual Cost. SPI = EV/PV, CPI = EV/AC, "
                         "EAC = AC + (BAC - EV)/CPI."),
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
        ("Predecessors", "Enter the predecessor's UID. A predecessor must be a detail task listed ABOVE its successor "
                         "(order tasks roughly by date). Links that break this, or point to an empty row or a missing UID, are "
                         "ignored and explained in the Link Check column (red). Summary rows ignore Duration / Predecessors / "
                         "Actuals - they roll up. The rule keeps the schedule and the critical path consistent and rules out "
                         "circular references."),
        ("Mode", "Auto = scheduled from predecessors, project start and the constraint. Manual = pinned to the "
                 "Constraint Date."),
        ("Constraints", "Pick an MS Project constraint and enter its Constraint Date (a date alone = Start No Earlier Than). "
                        "Start / Finish No Earlier Than push the task later; Start / Finish No Later Than and Must Start / "
                        "Finish On win over predecessors (as in MS Project), so a conflict shows as negative slack. As Late "
                        "As Possible moves the task later by its free slack: it finishes just in time for its successors "
                        "and moves nothing else (its planned dates move by the same days)."),
        ("Rescheduling", "Settings > RescheduleAfterStatus = Yes moves incomplete work to after the status date (like MS "
                         "Project 'Update Project'), so late tasks push successors."),
        ("Capacity", f"Tables grow when you add rows: Tasks starts with {sizes.tasks} rows, Resources {sizes.resources}, "
                     f"SprintTable {sizes.sprints}, Log 150, Holidays 100. The Gantt shows the first {sizes.tasks} "
                     f"tasks (the Gantt says so when there are more); the Gantt timeline has {sizes.gantt_columns} columns; "
                     f"the per-resource grids on the Resource Report show the first {sizes.grid_resources} "
                     f"resources and the Resource Report {sizes.weeks} weeks. To change these sizes, rebuild with "
                     "scripts/build_planner.py --tasks / --resources / --grid-resources / --sprints / "
                     "--gantt-columns / --weeks (see README)."),
        ("Speed", "An edit recalculates in well under a second for typical plans (about 0.05 s for 200 tasks). Each change "
                  "to the build is timed in CI; see docs/ARCHITECTURE.md in the repository. Keep the status date typed "
                  "(not =TODAY()), which would recalculate everything after every edit."),
        ("Views not updating?", "Formulas > Calculation Options must be Automatic. Excel takes the mode from the first "
                                "workbook opened in a session, so another workbook in Manual mode makes this one manual "
                                "too: switch back to Automatic, or press F9 to recalculate."),
        ("MOVING TO MICROSOFT PROJECT", None),
        ("Best: MS Project XML", "Save this workbook, then convert it with the browser page docs/export.html in the "
                                 "repository (nothing to install; works on a phone; the file stays on your device) or "
                                 "python scripts/export_mspdi.py. In MS Project: File > Open > the .xml > 'Open as a new "
                                 "project'. It carries the calendar and holidays, links by UID with lags, constraints, "
                                 "deadlines, actuals, the baseline, resources, assignments, Sprint (Text1) and Story "
                                 "Points (Number1), so MS Project schedules the plan the same way."),
        ("From MS Project", "MS Project > File > Save As > XML, then choose the .xml on docs/export.html (or run "
                            "python scripts/import_mspdi.py): you get a new planner workbook with the project's tasks, "
                            "links, calendar, holidays, resources, assignments, baseline and leave. Anything the planner "
                            "can't hold is listed and noted on the task."),
        ("Upgrading", "To move a plan made with an older version of this template into the latest one, choose the old "
                      "workbook on docs/export.html with 'Move to the latest template' (or run python "
                      "scripts/upgrade_planner.py old.xlsx new.xlsx). Everything you typed is carried over."),
    ] + (WIZARD if wizard else []) + [
        ("NAMES AND TABLES (Formulas > Name Manager)", None),
        ("Tables", "Tasks, ResourceTable, SprintTable, HolidayTable, GanttView, KPIs, StatusSummary, "
                   "NextMilestones, OverallStatus, StatusRAG, EarnedValue, StatusNotes, PhaseSummary, MilestoneReport, "
                   "LateTasks, StartingSoon, DueSoon, CriticalTasks, ResourceSummary, AgileKPIs, SprintBacklog, SprintBurndown, "
                   "Log, TopRAID, ChangeControl, SnapshotNow, StatusHistory and the List* "
                   "dropdown tables. Inside a table, formulas read [@Column] (this row)."),
        ("Settings", "ProjectStart, StatusDate, StatusWorkday, WorkWeek, Holidays, HoursPerDay, DefaultUnits, "
                     "RescheduleAfterStatus, GanttStart, GanttPeriod, OutlineShowLevel, RoadmapStart, Sprint1Start, SprintLength, SprintCount, SprintPrefix, "
                     "SprintLastEnd, CriticalSlackDays, ProjectFinish, PlannedFinish, FinishVarianceDays, LookaheadDays, "
                     "RAGAmberSlipDays, RAGRedSlipDays, RAGAmberIndex, RAGRedIndex, ResourceWeekStart, CurrentSprint, TotalPoints, "
                     "RAIDHighScore, RAIDMediumScore, NextFreeUID. Report values: ProjectRAG, "
                     "evm_BAC, evm_PV, evm_EV, evm_AC, evm_SPI, evm_CPI, evm_EAC ..."),
        ("Task columns", "tsk_ID, tsk_UID, tsk_WBS, tsk_Name, tsk_Duration, tsk_PlanStart, tsk_FcFinish, tsk_Status, tsk_Resource, "
                         "tsk_Work, tsk_Cost ... each is Tasks[Column], so it grows with the table. Other sheets read Tasks "
                         "only through these names."),
        ("Resources / sprints / Log", "res_Name, res_MaxUnits, res_Rate, ...  spr_Name, spr_Start, spr_End, spr_WorkDays, "
                                      "spr_Capacity, spr_Work ...  log_Type, log_Status, log_Rating, log_Open ..."),
        ("Key formulas", "Calculated start = WORKDAY.INTL(MIN(MAX(ProjectStart, [@_ConstraintLow], predecessor drivers), "
                         "[@_ConstraintHigh]) - 1, 1, WorkWeek, Holidays).  Predecessor driver (FS) = WORKDAY.INTL(OFFSET(Tasks[[#Headers],[_LeafCalcFinish]], "
                         "[@_Pred1Pos], 0), 1 + Lag, WorkWeek, Holidays).  Planned Start = Baseline Start if set, "
                         "else calculated.  Summary start = MIN(INDEX(Tasks[_LeafCalcStart], [@_Row]+1):INDEX(Tasks[_LeafCalcStart], "
                         "[@_Row]+[@_Descendants])), where _Descendants = COUNTIF(Tasks[WBS], [@WBS] & \".*\")."),
    ]


# MS Project's Import Wizard: only with the MSP_Tasks / MSP_Resources sheets (build option --import-wizard-sheets)
WIZARD = [
        ("Or: Import Wizard", "The MSP_Tasks / MSP_Resources sheets. Carries less than the XML (no calendar or holidays)."),
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
]


def build_instructions(wb, sizes, wizard=False):
    ws = wb[START]                                  # created by build_settings, which wrote the essential settings
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 110
    ws["B1"] = "Project Planner - start here"
    ws["B1"].font = font(size=18, bold=True, color="1F4E78")
    ws["B2"] = "Blue = type here. Grey = calculated. Every sheet works in Excel alone: desktop, web and iPhone."
    ws["B2"].font = NOTE_FONT
    _header(ws, START_ROW - 1, "YOUR PROJECT")
    r = START_ROW + len(ESSENTIAL) + 1
    for heading, steps in (("QUICK START", QUICK_START), ("EVERY WEEK", WEEKLY)):
        _header(ws, r, heading)
        r += 1
        for label, body, sheet in steps:
            ws.cell(r, 2, label).font = font(bold=True)
            c = ws.cell(r, 3, body + (f"  Go to {sheet}." if sheet else ""))
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if sheet:                               # a link inside the workbook (a hyperlink, not a formula)
                c.hyperlink = Hyperlink(ref=c.coordinate, location=f"'{sheet}'!A1")
                c.font = font(color="1F4E78", underline="single")
            r += 1
        r += 1
    ws.freeze_panes = "A3"
    _header(ws, r, "REFERENCE")
    r += 1
    for label, body in text(sizes, wizard):
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


def _header(ws, r, label):
    ws.cell(r, 2, label).font = HEADER_FONT
    ws.cell(r, 2).fill = HEADER_FILLS["in"]
    ws.cell(r, 3).fill = HEADER_FILLS["in"]
