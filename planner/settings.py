"""Settings (one named cell per setting), dropdown lists and the holiday calendar."""
import datetime as dt

from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation

from .core import (BOX, CALC_FILL, CALC_FONT, DATE, INPUT_FONT, NOTE_FONT, Col, add_table, fill, font, header_cell,
                   name, title)

LISTS = {  # table name, column name, named range, values
    "L_DepType": ("ListDepType", "Link Type", ["FS", "SS", "FF", "SF"]),
    "L_Mode": ("ListMode", "Task Mode", ["Auto", "Manual"]),
    "L_YesNo": ("ListYesNo", "Yes / No", ["Yes", "No"]),
    "L_ResType": ("ListResType", "Resource Type", ["Work", "Material", "Cost"]),
    "L_Period": ("ListPeriod", "Gantt Days per Column", [1, 7]),
    "L_RAG": ("ListRAG", "RAG", ["Green", "Amber", "Red"]),
    "L_RAIDType": ("ListRAIDType", "RAID Type", ["Risk", "Assumption", "Issue", "Dependency"]),
    "L_RAIDStatus": ("ListRAIDStatus", "RAID Status", ["Open", "In Progress", "Closed"]),
    "L_DecisionStatus": ("ListDecisionStatus", "Decision Status", ["Proposed", "Approved", "Rejected", "Superseded"]),
    "L_ChangeStatus": ("ListChangeStatus", "Change Status", ["Submitted", "Under Review", "Approved", "Rejected",
                                                             "Implemented"]),
}


def settings_rows(blank, max_sprints):
    """(name, label, value, number format, is input, note). Values that are formulas reference other names only."""
    work_week = "&".join(f'IF({d}="Yes","0","1")' for d in
                         ("WorkMon", "WorkTue", "WorkWed", "WorkThu", "WorkFri", "WorkSat", "WorkSun"))
    return [
        ("ProjectName", "Project name", "New Project" if blank else "Customer Portal Implementation", "@", True,
         "Shown on every sheet title."),
        ("ProjectManager", "Project manager", None if blank else "Priya Sharma", "@", True, ""),
        ("ProjectStartDate", "Project start date", None if blank else dt.date(2026, 9, 7), DATE, True,
         "Required. Tasks without predecessors start here (moved to the next working day if needed)."),
        ("ProjectStart", "Project start (in use)", "=MAX(1,N(ProjectStartDate))", DATE, False,
         "Copy of the date above (dates show as 1900 until it is entered)."),
        ("StatusDateInput", "Status date", None if blank else dt.date(2026, 10, 25), DATE, True,
         "The date progress is reported as of. Type it, or press Ctrl+; for today. It is a typed date rather than "
         "TODAY() on purpose: TODAY() would make Excel recalculate the whole schedule after every edit."),
        ("StatusDate", "Status date (in use)", "=IF(N(StatusDateInput)>0,StatusDateInput,ProjectStart)", DATE, False,
         "The project start until a status date is entered. Drives Status, the red status-date line on the Gantt "
         "and rescheduling of late work."),
        ("StatusWorkday", "First working day on/after status date", "=WORKDAY.INTL(StatusDate-1,1,WorkWeek,Holidays)",
         DATE, False, ""),
        ("RescheduleAfterStatus", "Reschedule incomplete work after status date?", "Yes", "@", True,
         "Yes = like MS Project 'move incomplete parts after status date': unstarted tasks can't start before the status "
         "date and in-progress tasks finish their remaining duration after it, pushing all successors."),
        ("HoursPerDay", "Working hours per day", 8, "0.0", True, "Work (hrs) = Duration x Hours/day x Units."),
        ("DefaultUnits", "Default assignment units", 1, "0%", True,
         "Used when a task has a resource but Units is blank. 100% = full time."),
        ("WorkMon", "Monday is a working day", "Yes", "@", True,
         "Working-week calendar (Yes/No). Feeds every WORKDAY / NETWORKDAYS formula."),
        ("WorkTue", "Tuesday is a working day", "Yes", "@", True, ""),
        ("WorkWed", "Wednesday is a working day", "Yes", "@", True, ""),
        ("WorkThu", "Thursday is a working day", "Yes", "@", True, ""),
        ("WorkFri", "Friday is a working day", "Yes", "@", True, ""),
        ("WorkSat", "Saturday is a working day", "No", "@", True, ""),
        ("WorkSun", "Sunday is a working day", "No", "@", True, ""),
        ("WorkWeek", "Weekend mask (calculated)", f"={work_week}", "@", False,
         "7-character Mon..Sun string used by WORKDAY.INTL (1 = non-working)."),
        ("GanttStart", "Gantt timeline start", "=ProjectStart-WEEKDAY(ProjectStart,3)", DATE, True,
         "Defaults to the Monday of the project start week. Overwrite with any date to scroll the timeline."),
        ("GanttPeriod", "Gantt days per column", 1, "0", True, "1 = daily timeline, 7 = weekly timeline."),
        ("Sprint1Start", "Sprint 1 start date", "=ProjectStart", DATE, True,
         "Defaults to the project start. Overwrite to align with your sprint calendar."),
        ("SprintLength", "Sprint length (calendar days)", 14, "0", True, "14 = two-week sprints."),
        ("SprintCount", "Number of sprints", 12, "0", True,
         f"Up to the number of rows in the Sprints table ({max_sprints}; add rows to the table for more). "
         "Tasks outside the sprint range are labelled 'Backlog'."),
        ("SprintPrefix", "Sprint name prefix", "Sprint ", "@", True, "Sprint names = prefix & number."),
        ("SprintLastEnd", "Last sprint end (calculated)", "=INDEX(spr_End,SprintCount)", DATE, False, ""),
        ("CurrentSprint", "Sprint containing the status date (calculated)",
         '=IF(OR(StatusDate<Sprint1Start,StatusDate>SprintLastEnd),"",INDEX(spr_Name,MATCH(StatusDate,spr_Start,1)))',
         "@", False, "Blank outside the sprint calendar. Agile Report shows its backlog and daily burndown."),
        ("CurrentSprintStart", "Current sprint start (calculated)",
         '=IF(CurrentSprint="","",INDEX(spr_Start,MATCH(CurrentSprint,spr_Name,0)))', DATE, False, ""),
        ("CurrentSprintEnd", "Current sprint end (calculated)",
         '=IF(CurrentSprint="","",INDEX(spr_End,MATCH(CurrentSprint,spr_Name,0)))', DATE, False, ""),
        ("TotalPoints", "Total story points (calculated)", "=SUM(tsk_LeafPoints)", "0.#", False,
         "Story points on all detail tasks: the release scope for burndown and burn-up."),
        ("CriticalSlackDays", "Critical when total slack is at most (days)", 0, "0", True,
         "0 = only tasks with no slack (or negative slack from a missed deadline) are critical, like MS Project."),
        ("ProjectFinish", "Forecast project finish (calculated)", "=MAX(tsk_FcFinish)", DATE, False,
         "The critical path's backward pass starts here."),
        ("PlannedFinish", "Planned project finish (calculated)", "=MAX(tsk_PlanFinish)", DATE, False,
         "Baseline finish once a baseline is set."),
        ("FinishVarianceDays", "Project finish variance (calculated, working days)",
         "=IF(ProjectFinish>=PlannedFinish,NETWORKDAYS.INTL(PlannedFinish,ProjectFinish,WorkWeek,Holidays)-1,"
         "1-NETWORKDAYS.INTL(ProjectFinish,PlannedFinish,WorkWeek,Holidays))", "+0;-0;0", False,
         "Forecast finish minus planned finish. Positive = late."),
        ("LookaheadDays", "Look-ahead window (working days)", 10, "0", True,
         "Schedule Report lists tasks starting or due within this many working days of the status date."),
        ("LookaheadEnd", "Look-ahead window ends (calculated)", "=WORKDAY.INTL(StatusDate-1,LookaheadDays,WorkWeek,Holidays)",
         DATE, False, ""),
        ("RAGAmberSlipDays", "RAG: amber when the finish slips by (working days)", 3, "0", True,
         "Status Report thresholds. A finish slip of at least this many working days is amber."),
        ("RAGRedSlipDays", "RAG: red when the finish slips by (working days)", 10, "0", True, ""),
        ("RAGAmberIndex", "RAG: amber when SPI or CPI is below", 0.95, "0.00", True,
         "SPI = earned value / planned value; CPI = earned value / actual cost. 1.00 = on plan."),
        ("RAGRedIndex", "RAG: red when SPI or CPI is below", 0.85, "0.00", True, ""),
        ("RAIDHighScore", "RAID: high rating from score", 15, "0", True,
         "Score = probability x impact (1-25). At or above this a risk or issue is rated High."),
        ("RAIDMediumScore", "RAID: medium rating from score", 8, "0", True, ""),
        ("ResourceWeekStart", "Resource Report first week", "=MAX(2,StatusDate-WEEKDAY(StatusDate,3)-14)", DATE, True,
         "Defaults to the Monday two weeks before the status date. Overwrite with any Monday to scroll the weekly grids."),
        ("NextFreeUID", "Next free task UID (calculated)", "=MAX(tsk_UID)+1", "0", False,
         "Give a newly inserted task row this UID (see Instructions > Inserting tasks)."),
    ]


def build_settings(wb, blank, max_sprints):
    st = wb.create_sheet("Settings")
    title(st, "Project Settings",
          "Blue = edit me. Grey = calculated. Column D shows the defined name every formula uses for that setting.")
    for col, (text, width) in enumerate([("Setting", 34), ("Value", 24), ("Named range", 24), ("Notes", 80)], start=2):
        header_cell(st.cell(4, col), text)
        st.column_dimensions[chr(64 + col)].width = width
    st.column_dimensions["A"].width = 3
    rows = {}
    for i, (n, label, value, fmt, is_input, note) in enumerate(settings_rows(blank, max_sprints)):
        r = 5 + i
        rows[n] = r
        st.cell(r, 2, label).font = font()
        c = st.cell(r, 3, value)
        c.number_format = fmt
        c.font = INPUT_FONT if is_input else CALC_FONT
        if not is_input:
            c.fill = CALC_FILL
        c.border = BOX
        st.cell(r, 4, n).font = font(color="7F7F7F")
        st.cell(r, 5, note).font = NOTE_FONT
        name(wb, n, f"Settings!$C${r}")

    yes_no = DataValidation(type="list", formula1="=L_YesNo", allow_blank=False)
    st.add_data_validation(yes_no)
    for n in ("RescheduleAfterStatus", "WorkMon", "WorkTue", "WorkWed", "WorkThu", "WorkFri", "WorkSat", "WorkSun"):
        yes_no.add(f"C{rows[n]}")
    period = DataValidation(type="list", formula1="=L_Period")
    st.add_data_validation(period)
    period.add(f"C{rows['GanttPeriod']}")
    sprints = DataValidation(type="whole", operator="between", formula1="1", formula2=str(max_sprints))
    st.add_data_validation(sprints)
    sprints.add(f"C{rows['SprintCount']}")
    for n in ("ProjectStartDate", "StatusDateInput"):
        cell = f"C{rows[n]}"
        st.conditional_formatting.add(cell, FormulaRule(formula=[f"N({cell})=0"], fill=fill("FFFF00")))
    return st


def build_lists(wb):
    ws = wb.create_sheet("Lists")
    for i, (n, (table, column, values)) in enumerate(LISTS.items()):
        col = 1 + i * 2
        add_table(ws, table, 1, col, [Col(column, "in", 16, values=values)], len(values), autofilter=False)
        name(wb, n, f"{table}[{column}]")
    ws.cell(len(max((v[2] for v in LISTS.values()), key=len)) + 3, 1,
            "Dropdown sources. Add a value by typing it in the row under a list: the table and its named range grow."
            ).font = NOTE_FONT
    return ws


HOLIDAYS = [(dt.date(2026, 12, 24), "Christmas Eve (example - replace with your calendar)"),
            (dt.date(2026, 12, 25), "Christmas Day (example)"),
            (dt.date(2026, 12, 31), "New Year's Eve (example)"),
            (dt.date(2027, 1, 1), "New Year's Day (example)")]


def build_holidays(wb, blank, n_rows=100):
    ws = wb.create_sheet("Holidays")
    title(ws, "Holidays & Non-Working Days",
          "Every date in this table is skipped by all scheduling formulas (named range: Holidays). Add rows to the table as needed.")
    data = [] if blank else HOLIDAYS
    add_table(ws, "HolidayTable", 4, 2, [
        Col("Date", "in", 18, DATE, values=[d for d, _ in data]),
        Col("Description", "in", 50, values=[t for _, t in data]),
    ], n_rows)
    name(wb, "Holidays", "HolidayTable[Date]")
    return ws
