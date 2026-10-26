"""Build an MS-Project-style planning workbook (Tasks, Gantt, Resources, Sprints, MSP export)."""
import argparse
import datetime as dt
import shutil
import subprocess
import tempfile
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import FormulaRule, ColorScaleRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as CL
from openpyxl.comments import Comment
from openpyxl.chart import BarChart, Reference

ARGS = argparse.ArgumentParser(description=__doc__)
ARGS.add_argument("out", help="output .xlsx path")
ARGS.add_argument("--blank", action="store_true", help="empty template: no sample tasks, resources, holidays")
ARGS.add_argument("--recalc", action="store_true",
                  help="recalculate with LibreOffice so the saved file carries values (for previews)")
ARGS = ARGS.parse_args()
OUT = ARGS.out
BLANK = ARGS.blank
N_TASKS = 200          # task rows available
FIRST = 6              # first data row on Tasks / Gantt
LAST = FIRST + N_TASKS - 1
N_RES = 50
N_SPR = 26
N_GANTT = 150          # timeline columns

WB = openpyxl.Workbook()
FONT = "Arial"


def F(**k):
    k.setdefault("name", FONT)
    k.setdefault("size", 10)
    return Font(**k)


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


INPUT_FONT = F(color="0000FF")
CALC_FILL = fill("F2F2F2")
HDR_IN = fill("1F4E78")
HDR_CALC = fill("595959")
HDR_HELP = fill("A6A6A6")
WHITE_B = F(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="D9D9D9")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
DATE = "ddd dd-mmm-yy"
WRAP_C = Alignment(horizontal="center", vertical="center", wrap_text=True)


def name(n, ref):
    WB.defined_names[n] = DefinedName(n, attr_text=ref)


def title(ws, text, sub):
    ws["A1"] = text
    ws["A1"].font = F(size=16, bold=True, color="1F4E78")
    ws["A2"] = sub
    ws["A2"].font = F(italic=True, color="595959")


def header(ws, row, col, text, kind="in", width=None):
    c = ws.cell(row, col, text)
    c.font = WHITE_B
    c.fill = {"in": HDR_IN, "calc": HDR_CALC, "help": HDR_HELP}[kind]
    c.alignment = WRAP_C
    c.border = BOX
    if width:
        ws.column_dimensions[CL(col)].width = width
    return c


# ---------------------------------------------------------------- Lists
lists = WB.active
lists.title = "Lists"
LISTS = {
    "L_DepType": ["FS", "SS", "FF", "SF"],
    "L_Mode": ["Auto", "Manual"],
    "L_YesNo": ["Yes", "No"],
    "L_ResType": ["Work", "Material", "Cost"],
    "L_Period": [1, 7],
}
for i, (n, vals) in enumerate(LISTS.items()):
    col = i + 1
    header(lists, 1, col, n, "calc", 14)
    for r, v in enumerate(vals, start=2):
        lists.cell(r, col, v).font = F()
    name(n, f"Lists!${CL(col)}$2:${CL(col)}${len(vals) + 1}")
lists.cell(8, 1, "Dropdown sources. Add values to a list and extend its named range (Formulas > Name Manager).").font = F(italic=True)

# ---------------------------------------------------------------- Settings
st = WB.create_sheet("Settings", 0)
title(st, "Project Settings", "Blue = edit me. Grey = calculated. Column D shows the named range every formula uses.")
for col, (t, w) in enumerate([("Setting", 34), ("Value", 24), ("Named range", 24), ("Notes", 80)], start=2):
    header(st, 4, col, t, "calc", w)
st.column_dimensions["A"].width = 3

SETTINGS = [
    # (name, label, value, number_format, is_input, note)
    ("ProjectName", "Project name", "New Project" if BLANK else "Customer Portal Implementation", "@", True, "Shown on every sheet title."),
    ("ProjectManager", "Project manager", None if BLANK else "Priya Sharma", "@", True, ""),
    ("ProjectStartDate", "Project start date", None if BLANK else dt.date(2026, 9, 7), DATE, True, "Tasks without predecessors start here (moved to the next working day if needed). Set this before entering tasks."),
    ("ProjectStart", "Project start (in use)", "=IF(N(ProjectStartDate)=0,TODAY(),ProjectStartDate)", DATE, False, "Today until a start date is entered above."),
    ("StatusDateOverride", "Status date override", None, DATE, True, "Leave blank to use TODAY(). Type a date to report progress 'as of' that day."),
    ("StatusDate", "Status date (in use)", "=IF(StatusDateOverride=\"\",TODAY(),StatusDateOverride)", DATE, False, "Drives Status, the red 'today' line on the Gantt and rescheduling of late work."),
    ("StatusWorkday", "First working day on/after status date", "=WORKDAY.INTL(StatusDate-1,1,WorkWeek,Holidays)", DATE, False, ""),
    ("RescheduleAfterStatus", "Reschedule incomplete work after status date?", "Yes", "@", True, "Yes = like MS Project 'move incomplete parts after status date': unstarted tasks can't start before the status date and in-progress tasks finish their remaining duration after it, pushing all successors."),
    ("HoursPerDay", "Working hours per day", 8, "0.0", True, "Work (hrs) = Duration x Hours/day x Units."),
    ("DefaultUnits", "Default assignment units", 1, "0%", True, "Used when a task has a resource but Units is blank. 100% = full time."),
    ("WorkMon", "Monday is a working day", "Yes", "@", True, "Working-week calendar (Yes/No). Feeds every WORKDAY / NETWORKDAYS formula."),
    ("WorkTue", "Tuesday is a working day", "Yes", "@", True, ""),
    ("WorkWed", "Wednesday is a working day", "Yes", "@", True, ""),
    ("WorkThu", "Thursday is a working day", "Yes", "@", True, ""),
    ("WorkFri", "Friday is a working day", "Yes", "@", True, ""),
    ("WorkSat", "Saturday is a working day", "No", "@", True, ""),
    ("WorkSun", "Sunday is a working day", "No", "@", True, ""),
    ("WorkWeek", "Weekend mask (calculated)", '=IF(WorkMon="Yes","0","1")&IF(WorkTue="Yes","0","1")&IF(WorkWed="Yes","0","1")&IF(WorkThu="Yes","0","1")&IF(WorkFri="Yes","0","1")&IF(WorkSat="Yes","0","1")&IF(WorkSun="Yes","0","1")', "@", False, "7-character Mon..Sun string used by WORKDAY.INTL (1 = non-working)."),
    ("GanttStart", "Gantt timeline start", "=ProjectStart-WEEKDAY(ProjectStart,3)", DATE, True, "Defaults to the Monday of the project start week. Overwrite with any date to scroll the timeline."),
    ("GanttPeriod", "Gantt days per column", 1, "0", True, "1 = daily timeline, 7 = weekly timeline (150 columns either way)."),
    ("Sprint1Start", "Sprint 1 start date", "=ProjectStart", DATE, True, "Defaults to the project start. Overwrite to align with your sprint calendar."),
    ("SprintLength", "Sprint length (calendar days)", 14, "0", True, "14 = two-week sprints."),
    ("SprintCount", "Number of sprints", 12, "0", True, f"Up to {N_SPR}. Tasks outside the sprint range are labelled 'Backlog'."),
    ("SprintPrefix", "Sprint name prefix", "Sprint ", "@", True, "Sprint names = prefix & number."),
    ("SprintLastEnd", "Last sprint end (calculated)", "=INDEX(spr_End,SprintCount)", DATE, False, ""),
    ("NextFreeUID", "Next free task UID (calculated)", "=MAX(tsk_UID)+1", "0", False, "Give a newly inserted task row this UID (see Instructions > Inserting tasks)."),
]
for i, (n, label, val, fmt, is_in, note) in enumerate(SETTINGS):
    r = 5 + i
    st.cell(r, 2, label).font = F()
    c = st.cell(r, 3, val)
    c.number_format = fmt
    c.font = INPUT_FONT if is_in else F()
    if not is_in:
        c.fill = CALC_FILL
    c.border = BOX
    st.cell(r, 4, n).font = F(color="7F7F7F")
    st.cell(r, 5, note).font = F(italic=True, color="595959")
    name(n, f"Settings!$C${r}")
SROW = {n: 5 + i for i, (n, *_r) in enumerate(SETTINGS)}
yn = DataValidation(type="list", formula1="=L_YesNo", allow_blank=False)
st.add_data_validation(yn)
for n in ["RescheduleAfterStatus", "WorkMon", "WorkTue", "WorkWed", "WorkThu", "WorkFri", "WorkSat", "WorkSun"]:
    yn.add(f"C{SROW[n]}")
dvp = DataValidation(type="list", formula1="=L_Period")
st.add_data_validation(dvp)
dvp.add(f"C{SROW['GanttPeriod']}")
dvs = DataValidation(type="whole", operator="between", formula1="1", formula2=str(N_SPR))
st.add_data_validation(dvs)
dvs.add(f"C{SROW['SprintCount']}")
st.conditional_formatting.add(f"C{SROW['ProjectStartDate']}", FormulaRule(formula=[f"N(C{SROW['ProjectStartDate']})=0"], fill=fill("FFFF00")))

# ---------------------------------------------------------------- Holidays
ho = WB.create_sheet("Holidays")
title(ho, "Holidays & Non-Working Days", "Every date listed here is skipped by all scheduling formulas (named range: Holidays).")
header(ho, 4, 2, "Date", "in", 18)
header(ho, 4, 3, "Description", "in", 40)
HOL = [] if BLANK else [(dt.date(2026, 12, 24), "Christmas Eve (example - replace with your calendar)"),
       (dt.date(2026, 12, 25), "Christmas Day (example)"),
       (dt.date(2026, 12, 31), "New Year's Eve (example)"),
       (dt.date(2027, 1, 1), "New Year's Day (example)")]
for r in range(5, 105):
    for c in (2, 3):
        ho.cell(r, c).font = INPUT_FONT
        ho.cell(r, c).border = BOX
    ho.cell(r, 2).number_format = DATE
for i, (d, t) in enumerate(HOL):
    ho.cell(5 + i, 2, d)
    ho.cell(5 + i, 3, t)
name("Holidays", "Holidays!$B$5:$B$104")

# ---------------------------------------------------------------- Resources
rs = WB.create_sheet("Resources")
title(rs, "Resource Sheet", "Like MS Project's Resource Sheet. Names here feed the Resource dropdown on Tasks.")
RCOLS = [("Resource Name", "in", 22), ("Initials", "in", 9), ("Group / Role", "in", 20), ("Type", "in", 10),
         ("Max Units", "in", 10), ("Std Rate / hr", "in", 12), ("Email", "in", 26),
         ("# Tasks", "calc", 9), ("Assigned Work (hrs)", "calc", 13), ("Assigned Cost", "calc", 13), ("Notes", "in", 30)]
rs.column_dimensions["A"].width = 3
for i, (t, k, w) in enumerate(RCOLS):
    header(rs, 5, 2 + i, t, k, w)
RES = [] if BLANK else [("Priya Sharma", "PS", "Project Manager", 0.5, 95, "priya@example.com"),
       ("Morgan Patel", "MP", "Business Analyst", 1, 70, "morgan@example.com"),
       ("Alex Chen", "AC", "Solution Architect", 0.5, 110, "alex@example.com"),
       ("Casey Kim", "CK", "UX Designer", 1, 75, "casey@example.com"),
       ("Jordan Lee", "JL", "Backend Developer", 1, 85, "jordan@example.com"),
       ("Sam Rivera", "SR", "Frontend Developer", 1, 80, "sam@example.com"),
       ("Taylor Morgan", "TM", "QA Engineer", 1, 65, "taylor@example.com")]
R0 = 6
for i in range(N_RES):
    r = R0 + i
    for c in range(2, 13):
        rs.cell(r, c).border = BOX
        rs.cell(r, c).font = INPUT_FONT
    if i < len(RES):
        nm, ini, role, units, rate, mail = RES[i]
        for c, v in zip(range(2, 9), [nm, ini, role, "Work", units, rate, mail]):
            rs.cell(r, c, v)
    rs.cell(r, 6).number_format = "0%"
    rs.cell(r, 7).number_format = "#,##0.00"
    rs.cell(r, 9, f'=IF($B{r}="","",COUNTIFS(tsk_Resource,$B{r},tsk_IsLeaf,1))')
    rs.cell(r, 10, f'=IF($B{r}="","",SUMIFS(tsk_Work,tsk_Resource,$B{r},tsk_IsLeaf,1))')
    rs.cell(r, 11, f'=IF($B{r}="","",SUMIFS(tsk_Cost,tsk_Resource,$B{r},tsk_IsLeaf,1))')
    for c, fmt in ((9, "0"), (10, "#,##0.0"), (11, "#,##0.00")):
        rs.cell(r, c).font = F()
        rs.cell(r, c).fill = CALC_FILL
        rs.cell(r, c).number_format = fmt
RL = R0 + N_RES - 1
for n, col in [("res_Name", "B"), ("res_Initials", "C"), ("res_Group", "D"), ("res_Type", "E"),
               ("res_MaxUnits", "F"), ("res_Rate", "G"), ("res_Email", "H")]:
    name(n, f"Resources!${col}${R0}:${col}${RL}")
dvt = DataValidation(type="list", formula1="=L_ResType")
rs.add_data_validation(dvt)
dvt.add(f"E{R0}:E{RL}")
rs.freeze_panes = "C6"

# ---------------------------------------------------------------- Tasks
ts = WB.create_sheet("Tasks", 1)
title(ts, "Task Entry (Schedule)", None)
ts["A1"] = '=ProjectName&"  -  Task Entry"'
ts["A2"] = ('Type in BLUE columns only. Level = outline level (1 = phase, 2 = task, 3 = sub-task). Predecessors refer to the UID column. '
            'Planned = Baseline once set, otherwise calculated from Duration + Predecessors. Forecast also uses Actuals, so slips ripple to every successor.')
ts["A2"].font = F(italic=True, color="595959")
ts["A3"] = '="Next free UID for a new row: "&NextFreeUID'
ts["A3"].font = F(bold=True, color="1F4E78")

COLS = [  # key, header, kind, width
    ("ID", "ID", "calc", 5), ("UID", "UID", "in", 6), ("WBS", "WBS", "calc", 8), ("Level", "Level", "in", 6),
    ("Name", "Task Name", "in", 36), ("Mode", "Mode", "in", 8), ("Dur", "Duration (work days)", "in", 9),
    ("P1", "Pred 1 UID", "in", 6), ("T1", "Type 1", "in", 6), ("L1", "Lag 1 (days)", "in", 6),
    ("P2", "Pred 2 UID", "in", 6), ("T2", "Type 2", "in", 6), ("L2", "Lag 2 (days)", "in", 6),
    ("Cons", "Start No Earlier Than / Manual Start", "in", 15), ("DL", "Deadline", "in", 15),
    ("PS", "Planned Start (baseline, else calculated)", "calc", 15), ("PF", "Planned Finish (baseline, else calculated)", "calc", 15),
    ("AS", "Actual Start", "in", 15), ("AF", "Actual Finish", "in", 15),
    ("FS", "Forecast Start", "calc", 15), ("FF", "Forecast Finish", "calc", 15),
    ("SVar", "Start Variance (work days)", "calc", 9), ("Var", "Finish Variance (work days)", "calc", 9),
    ("DLChk", "Deadline Check", "calc", 12), ("LinkChk", "Link Check", "calc", 24),
    ("Pct", "% Complete (enter)", "in", 9), ("Prog", "Progress % (rolled up)", "calc", 9),
    ("Status", "Status", "calc", 12), ("Res", "Resource", "in", 16), ("Units", "Units", "in", 7),
    ("Work", "Work (hrs)", "calc", 9), ("Cost", "Cost", "calc", 11), ("Sprint", "Sprint", "calc", 10),
    ("PredTxt", "Predecessors (by ID, MSP format)", "calc", 14), ("Notes", "Notes", "in", 40),
    # baseline: copy the 5 source columns and paste VALUES onto the 5 baseline columns
    ("SrcS", "Baseline source: Start", "calc", 15), ("SrcF", "Baseline source: Finish", "calc", 15),
    ("SrcD", "Baseline source: Duration", "calc", 9), ("SrcW", "Baseline source: Work (hrs)", "calc", 9),
    ("SrcC", "Baseline source: Cost", "calc", 11),
    ("BS", "Baseline Start", "in", 15), ("BF", "Baseline Finish", "in", 15), ("BD", "Baseline Duration", "in", 9),
    ("BW", "Baseline Work (hrs)", "in", 9), ("BC", "Baseline Cost", "in", 11),
    ("DVar", "Duration Variance (days)", "calc", 9), ("WVar", "Work Variance (hrs)", "calc", 9),
    ("CVar", "Cost Variance", "calc", 11),
    # helper columns (grouped & hidden)
    ("Lvl", "Lvl (effective)", "help", 7), ("Idx", "Row Index", "help", 7), ("Parent", "Parent Row", "help", 7),
    ("Desc", "Descendant Count", "help", 8), ("IsSum", "Is Summary", "help", 8), ("IsLeaf", "Is Leaf", "help", 7),
    ("CPS", "Calculated Plan Start", "help", 11), ("CPF", "Calculated Plan Finish", "help", 11),
    ("M1", "Pred1 Problem", "help", 14), ("M2", "Pred2 Problem", "help", 14),
    ("D1P", "Pred1 Plan Driver", "help", 11), ("D2P", "Pred2 Plan Driver", "help", 11),
    ("D1F", "Pred1 Fcst Driver", "help", 11), ("D2F", "Pred2 Fcst Driver", "help", 11),
    ("SN", "Fcst Start # (leaf)", "help", 11), ("FN", "Fcst Finish # (leaf)", "help", 11),
    ("WPD", "Work per Cal Day", "help", 9), ("WtDur", "Weight Dur", "help", 8), ("WtDone", "Weight Done", "help", 8),
    ("T1x", "Pred1 Text (ID)", "help", 9), ("T2x", "Pred2 Text (ID)", "help", 9),
    ("U1x", "Pred1 Text (UID)", "help", 9), ("U2x", "Pred2 Text (UID)", "help", 9),
    ("UPredTxt", "Unique ID Predecessors", "help", 12), ("MKey", "Open Milestone Key", "help", 10),
]
C = {}
for i, (k, h, kind, w) in enumerate(COLS):
    C[k] = CL(i + 1)
    header(ts, 5, i + 1, h, kind, w)
ts.row_dimensions[5].height = 66
# group labels row 4
for k, label in [("ID", "IDENTITY"), ("Mode", "SCHEDULING INPUTS"), ("PS", "PLAN / ACTUAL / FORECAST"),
                 ("SVar", "CHECKS"), ("Pct", "PROGRESS"), ("Res", "RESOURCES & SPRINT"),
                 ("SrcS", "SET BASELINE: copy these 5 columns, then Paste Special > Values onto the 5 Baseline columns"),
                 ("BS", ""), ("DVar", "VARIANCE vs BASELINE"), ("Lvl", "HELPER COLUMNS (hidden - click + above to show)")]:
    ts[f"{C[k]}4"] = label or None
    ts[f"{C[k]}4"].font = F(bold=True, color="1F4E78")

# header anchors for OFFSET look-ups (no circular references)
for k, n in [("ID", "hdr_ID"), ("WBS", "hdr_WBS"), ("CPS", "hdr_CalcStart"), ("CPF", "hdr_CalcFinish"),
             ("FS", "hdr_FcStart"), ("FF", "hdr_FcFinish"), ("Work", "hdr_Work"), ("Cost", "hdr_Cost"),
             ("IsLeaf", "hdr_IsLeaf"), ("WtDur", "hdr_WtDur"), ("WtDone", "hdr_WtDone")]:
    name(n, f"Tasks!${C[k]}$5")
TSK_NAMES = {
    "ID": "tsk_ID", "UID": "tsk_UID", "WBS": "tsk_WBS", "Lvl": "tsk_Level", "Name": "tsk_Name", "Mode": "tsk_Mode",
    "Dur": "tsk_Duration", "Cons": "tsk_Constraint", "DL": "tsk_Deadline", "PS": "tsk_PlanStart", "PF": "tsk_PlanFinish",
    "AS": "tsk_ActStart", "AF": "tsk_ActFinish", "FS": "tsk_FcStart", "FF": "tsk_FcFinish", "SVar": "tsk_StartVariance",
    "Var": "tsk_Variance", "DLChk": "tsk_DeadlineCheck", "LinkChk": "tsk_LinkCheck", "Pct": "tsk_PctInput",
    "Prog": "tsk_Progress", "Status": "tsk_Status", "Res": "tsk_Resource", "Units": "tsk_Units", "Work": "tsk_Work",
    "Cost": "tsk_Cost", "Sprint": "tsk_Sprint", "PredTxt": "tsk_Predecessors", "Notes": "tsk_Notes",
    "BS": "tsk_BaseStart", "BF": "tsk_BaseFinish", "BD": "tsk_BaseDuration", "BW": "tsk_BaseWork", "BC": "tsk_BaseCost",
    "DVar": "tsk_DurationVariance", "WVar": "tsk_WorkVariance", "CVar": "tsk_CostVariance",
    "IsSum": "tsk_IsSummary", "IsLeaf": "tsk_IsLeaf", "CPS": "tsk_CalcStart", "CPF": "tsk_CalcFinish",
    "SN": "tsk_StartN", "FN": "tsk_FinishN", "WPD": "tsk_WorkPerDay", "WtDur": "tsk_WtDur", "WtDone": "tsk_WtDone",
    "UPredTxt": "tsk_UIDPredecessors", "MKey": "tsk_MilestoneKey",
}
for k, n in TSK_NAMES.items():
    name(n, f"Tasks!${C[k]}${FIRST}:${C[k]}${LAST}")

WD = "WorkWeek,Holidays"


def wdiff(a, b):
    """Signed working-day difference b - a."""
    return f"IF({b}>={a},NETWORKDAYS.INTL({a},{b},{WD})-1,1-NETWORKDAYS.INTL({b},{a},{WD}))"


def link_problem(r, p):
    """Why a predecessor link can't be used ("" = fine). Uses no dates, so it can never be circular."""
    P, m = f"${C[p]}{r}", f"MATCH(${C[p]}{r},tsk_UID,0)"
    pw = f"INDEX(tsk_WBS,{m})"
    return (f'=IF(OR(N({P})=0,${C["Name"]}{r}=""),"",IF(${C["IsSum"]}{r},"ignored: summary tasks roll up",'
            f'IF(ISNA({m}),"UID "&{P}&" not found",IF({P}=${C["UID"]}{r},"links to itself",'
            f'IF(INDEX(tsk_Name,{m})="","UID "&{P}&" is an empty row",'
            f'IF(LEFT(${C["WBS"]}{r},LEN({pw})+1)={pw}&".","UID "&{P}&" is its own summary",""))))))')


def driver(r, p, t, l, mp, hs, hf):
    """Earliest start imposed by one predecessor link (FS/SS/FF/SF + lag), 0 if none or invalid."""
    P, T, Lg, D = f"${C[p]}{r}", f"${C[t]}{r}", f"N(${C[l]}{r})", f"${C['Dur']}{r}"
    m = f"MATCH({P},tsk_UID,0)"
    pf = f"OFFSET({hf},{m},0)"
    ps = f"OFFSET({hs},{m},0)"
    back = f"-MAX(0,N({D})-1)"
    return (f'=IF(OR(N({P})=0,${C[mp]}{r}<>"",${C["IsSum"]}{r},${C["Name"]}{r}=""),0,IFERROR(CHOOSE(MATCH(IF({T}="","FS",{T}),{{"FS","SS","FF","SF"}},0),'
            f'WORKDAY.INTL({pf},IF(N({D})=0,0,1)+{Lg},{WD}),'
            f'WORKDAY.INTL({ps},{Lg},{WD}),'
            f'WORKDAY.INTL(WORKDAY.INTL({pf},{Lg},{WD}),{back},{WD}),'
            f'WORKDAY.INTL(WORKDAY.INTL({ps},{Lg},{WD}),{back},{WD})),0))')


def pred_txt(r, p, t, l, mp, by_id):
    P, T, Lg = f"${C[p]}{r}", f'IF(${C[t]}{r}="","FS",${C[t]}{r})', f"N(${C[l]}{r})"
    ref = f"INDEX(tsk_ID,MATCH({P},tsk_UID,0))" if by_id else P
    return (f'=IF(OR(N({P})=0,${C[mp]}{r}<>""),"",{ref}&IF(AND({T}="FS",{Lg}=0),"",'
            f'{T}&IF({Lg}=0,"",IF({Lg}>0,"+","")&{Lg}&"d")))')


def block(hdr, r):
    return f"OFFSET({hdr},${C['Idx']}{r}+1,0,${C['Desc']}{r},1)"


def tasks_formulas(r):
    c = {k: f"${C[k]}{r}" for k in C}
    nm, s = c["Name"], c["IsSum"]
    f = {}
    f["ID"] = f'=IF({nm}="","",COUNTA(${C["Name"]}${FIRST}:{nm}))'
    f["Lvl"] = f'=IF({nm}="","",MAX(1,N({c["Level"]})))'
    f["Idx"] = f"=ROW()-ROW(hdr_ID)"
    f["Parent"] = (f'=IF(OR({nm}="",{c["Lvl"]}<=1),0,IFERROR(LOOKUP(2,1/(${C["Lvl"]}$5:${C["Lvl"]}{r-1}={c["Lvl"]}-1),'
                   f'ROW(${C["Lvl"]}$5:${C["Lvl"]}{r-1})),0))')
    f["WBS"] = (f'=IF({nm}="","",IF({c["Lvl"]}=1,COUNTIF(${C["Lvl"]}${FIRST}:{c["Lvl"]},1)&"",'
                f'IF({c["Parent"]}=0,"?",OFFSET(hdr_WBS,{c["Parent"]}-ROW(hdr_WBS),0)&"."&'
                f'COUNTIF(INDEX(${C["Lvl"]}:${C["Lvl"]},{c["Parent"]}+1):{c["Lvl"]},{c["Lvl"]}))))')
    f["Desc"] = f'=IF({nm}="",0,COUNTIF(tsk_WBS,{c["WBS"]}&".*"))'
    f["IsSum"] = f"={c['Desc']}>0"
    f["IsLeaf"] = f'=IF(AND({nm}<>"",NOT({s})),1,0)'
    f["M1"] = link_problem(r, "P1")
    f["M2"] = link_problem(r, "P2")
    f["LinkChk"] = (f'=IF({nm}="","",TRIM(IF({c["M1"]}<>"","Pred 1: "&{c["M1"]}&". ","")&'
                    f'IF({c["M2"]}<>"","Pred 2: "&{c["M2"]}&".","")))')
    f["D1P"] = driver(r, "P1", "T1", "L1", "M1", "hdr_CalcStart", "hdr_CalcFinish")
    f["D2P"] = driver(r, "P2", "T2", "L2", "M2", "hdr_CalcStart", "hdr_CalcFinish")
    f["D1F"] = driver(r, "P1", "T1", "L1", "M1", "hdr_FcStart", "hdr_FcFinish")
    f["D2F"] = driver(r, "P2", "T2", "L2", "M2", "hdr_FcStart", "hdr_FcFinish")
    manual = f'IF(N({c["Cons"]})>0,{c["Cons"]},ProjectStart)'
    f["CPS"] = (f'=IF({nm}="","",IF({s},MIN({block("hdr_CalcStart", r)}),IF({c["Mode"]}="Manual",{manual},'
                f'WORKDAY.INTL(MAX(ProjectStart,N({c["Cons"]}),{c["D1P"]},{c["D2P"]})-1,1,{WD}))))')
    f["CPF"] = (f'=IF({nm}="","",IF({s},MAX({block("hdr_CalcFinish", r)}),IF(N({c["Dur"]})<=0,{c["CPS"]},'
                f'WORKDAY.INTL({c["CPS"]},{c["Dur"]}-1,{WD}))))')
    f["PS"] = f'=IF({nm}="","",IF(N({c["BS"]})>0,{c["BS"]},{c["CPS"]}))'
    f["PF"] = f'=IF({nm}="","",IF(N({c["BF"]})>0,{c["BF"]},{c["CPF"]}))'
    f["FS"] = (f'=IF({nm}="","",IF({s},MIN({block("hdr_FcStart", r)}),IF(N({c["AS"]})>0,{c["AS"]},IF({c["Mode"]}="Manual",{manual},'
               f'WORKDAY.INTL(MAX(ProjectStart,N({c["Cons"]}),{c["D1F"]},{c["D2F"]},IF(RescheduleAfterStatus="Yes",StatusDate,0))-1,1,{WD})))))')
    natural = f'WORKDAY.INTL({c["FS"]},{c["Dur"]}-1,{WD})'
    remaining = f'WORKDAY.INTL(StatusWorkday,MAX(0,ROUNDUP({c["Dur"]}*(1-MIN(1,N({c["Pct"]}))),0)-1),{WD})'
    f["FF"] = (f'=IF({nm}="","",IF({s},MAX({block("hdr_FcFinish", r)}),IF(N({c["AF"]})>0,{c["AF"]},IF(N({c["Dur"]})<=0,{c["FS"]},'
               f'IF(AND(N({c["AS"]})>0,RescheduleAfterStatus="Yes"),MAX({natural},{remaining}),{natural})))))')
    f["SVar"] = f'=IF({nm}="","",{wdiff(c["PS"], c["FS"])})'
    f["Var"] = f'=IF({nm}="","",{wdiff(c["PF"], c["FF"])})'
    f["DLChk"] = (f'=IF(OR({nm}="",N({c["DL"]})=0),"",IF({c["FF"]}>{c["DL"]},'
                  f'"Late by "&({wdiff(c["DL"], c["FF"])})&"d","OK"))')
    f["Prog"] = (f'=IF({nm}="","",IF({s},IFERROR(SUM({block("hdr_WtDone", r)})/SUM({block("hdr_WtDur", r)}),0),'
                 f'IF(N({c["AF"]})>0,1,MIN(1,N({c["Pct"]})))))')
    f["Status"] = (f'=IF({nm}="","",IF({c["Prog"]}>=1,"Complete",IF(OR(N({c["AS"]})>0,{c["Prog"]}>0),'
                   f'IF({c["Var"]}>0,"Behind","In Progress"),IF({c["PS"]}<StatusDate,"Late Start","Not Started"))))')
    f["Work"] = (f'=IF({nm}="","",IF({s},SUMPRODUCT({block("hdr_Work", r)},{block("hdr_IsLeaf", r)}),'
                 f'IF({c["Res"]}="",0,N({c["Dur"]})*HoursPerDay*IF({c["Units"]}="",DefaultUnits,{c["Units"]}))))')
    f["Cost"] = (f'=IF({nm}="","",IF({s},SUMPRODUCT({block("hdr_Cost", r)},{block("hdr_IsLeaf", r)}),'
                 f'{c["Work"]}*IFERROR(INDEX(res_Rate,MATCH({c["Res"]},res_Name,0)),0)))')
    f["Sprint"] = (f'=IF(OR({nm}="",{s}),"",IF(OR({c["FS"]}<Sprint1Start,{c["FS"]}>SprintLastEnd),"Backlog",'
                   f'INDEX(spr_Name,MATCH({c["FS"]},spr_Start,1))))')
    f["T1x"] = pred_txt(r, "P1", "T1", "L1", "M1", True)
    f["T2x"] = pred_txt(r, "P2", "T2", "L2", "M2", True)
    f["U1x"] = pred_txt(r, "P1", "T1", "L1", "M1", False)
    f["U2x"] = pred_txt(r, "P2", "T2", "L2", "M2", False)
    join = lambda a, b: f'=IF({nm}="","",{c[a]}&IF(AND({c[a]}<>"",{c[b]}<>""),",","")&{c[b]})'
    f["PredTxt"] = join("T1x", "T2x")
    f["UPredTxt"] = join("U1x", "U2x")
    f["SrcS"] = f'=IF({nm}="","",{c["FS"]})'
    f["SrcF"] = f'=IF({nm}="","",{c["FF"]})'
    f["SrcD"] = f'=IF({nm}="","",IF({s},NETWORKDAYS.INTL({c["FS"]},{c["FF"]},{WD}),N({c["Dur"]})))'
    f["SrcW"] = f'=IF({nm}="","",{c["Work"]})'
    f["SrcC"] = f'=IF({nm}="","",{c["Cost"]})'
    var = lambda cur, base: f'=IF(OR({nm}="",{c[base]}=""),"",{c[cur]}-N({c[base]}))'
    f["DVar"] = var("SrcD", "BD")
    f["WVar"] = var("Work", "BW")
    f["CVar"] = var("Cost", "BC")
    f["SN"] = f'=IF({c["IsLeaf"]}=1,{c["FS"]},0)'
    f["FN"] = f'=IF({c["IsLeaf"]}=1,{c["FF"]},0)'
    f["WPD"] = f'=IF(AND({c["IsLeaf"]}=1,N({c["Work"]})>0),{c["Work"]}/({c["FN"]}-{c["SN"]}+1),0)'
    f["WtDur"] = f'=IF({c["IsLeaf"]}=1,MAX(N({c["Dur"]}),0.1),0)'
    f["MKey"] = f'=IF(AND({c["IsLeaf"]}=1,N({c["Dur"]})=0,{c["Prog"]}<1),{c["FF"]}+{c["Idx"]}/100000,"")'
    f["WtDone"] = f'=IF({c["IsLeaf"]}=1,{c["WtDur"]}*{c["Prog"]},0)'
    return f


D = dt.date
SAMPLE = [] if BLANK else [  # Level, Name, Mode, Dur, P1,T1,L1, P2,T2,L2, Cons, Deadline, ActStart, ActFinish, Pct, Res, Units, Notes
    (1, "Initiation", "Auto", None, None, None, None, None, None, None, None, None, None, None, None, None, None, "Phase 1 - summary rows roll up dates, work, cost and progress automatically."),
    (2, "Project charter & scope statement", "Auto", 3, None, None, None, None, None, None, None, None, D(2026, 9, 7), D(2026, 9, 10), 1, "Priya Sharma", None, "Took 1 extra day for sponsor sign-off."),
    (2, "Stakeholder kickoff workshop", "Auto", 2, 2, "FS", None, None, None, None, None, None, D(2026, 9, 11), D(2026, 9, 14), 1, "Morgan Patel", None, ""),
    (2, "Kickoff complete", "Auto", 0, 3, "FS", None, None, None, None, None, None, D(2026, 9, 14), D(2026, 9, 14), 1, None, None, "Milestone (Duration 0)."),
    (1, "Requirements & Design", "Auto", None, None, None, None, None, None, None, None, None, None, None, None, None, None, ""),
    (2, "Requirements gathering", "Auto", 8, 4, "FS", None, None, None, None, None, None, D(2026, 9, 15), D(2026, 9, 25), 1, "Morgan Patel", None, ""),
    (2, "Solution architecture", "Auto", 5, 6, "SS", 3, None, None, None, None, None, D(2026, 9, 18), D(2026, 9, 25), 1, "Alex Chen", None, "Start-to-Start + 3 day lag example."),
    (2, "UX wireframes & prototype", "Auto", 6, 6, "FS", None, None, None, None, None, None, D(2026, 9, 28), D(2026, 10, 6), 1, "Casey Kim", None, ""),
    (2, "Design sign-off", "Auto", 0, 7, "FS", None, 8, "FS", None, None, None, D(2026, 10, 6), D(2026, 10, 6), 1, None, None, "Two predecessors - waits for the later one."),
    (1, "Build", "Auto", None, None, None, None, None, None, None, None, None, None, None, None, None, None, ""),
    (2, "Backend", "Auto", None, None, None, None, None, None, None, None, None, None, None, None, None, None, "Level-2 summary with sub-tasks."),
    (3, "Data model & database setup", "Auto", 5, 9, "FS", None, None, None, None, None, None, D(2026, 10, 7), D(2026, 10, 14), 1, "Jordan Lee", None, ""),
    (3, "Core API endpoints", "Auto", 10, 12, "FS", None, None, None, None, None, None, D(2026, 10, 15), None, 0.5, "Jordan Lee", None, "In progress - enter % Complete; forecast finish adjusts."),
    (3, "Payment gateway integration", "Auto", 6, 13, "FS", None, None, None, None, None, None, None, None, None, "Jordan Lee", None, ""),
    (2, "Frontend", "Auto", None, None, None, None, None, None, None, None, None, None, None, None, None, None, ""),
    (3, "UI component library", "Auto", 6, 9, "FS", None, None, None, None, None, None, D(2026, 10, 7), D(2026, 10, 16), 1, "Sam Rivera", None, ""),
    (3, "Feature screens", "Auto", 12, 16, "FS", None, 13, "SS", 2, None, None, D(2026, 10, 19), None, 0.25, "Sam Rivera", None, ""),
    (3, "Accessibility review", "Auto", 3, 17, "FF", 1, None, None, None, None, None, None, None, None, "Casey Kim", 0.5, "Finish-to-Finish + 1 day; 50% units."),
    (2, "Build complete", "Auto", 0, 14, "FS", None, 18, "FS", None, None, D(2026, 11, 6), None, None, None, None, None, "Milestone with a deadline of 06-Nov (see Deadline Check)."),
    (1, "Test & Release", "Auto", None, None, None, None, None, None, None, None, None, None, None, None, None, None, ""),
    (2, "System integration testing", "Auto", 8, 19, "FS", None, None, None, None, None, None, None, None, None, "Taylor Morgan", None, ""),
    (2, "User acceptance testing", "Auto", 6, 21, "FS", None, None, None, None, D(2026, 12, 7), None, None, None, None, "Morgan Patel", None, "Start No Earlier Than 07-Dec (client availability)."),
    (2, "Training & user guides", "Auto", 5, 21, "SS", 2, None, None, None, None, None, None, None, None, "Casey Kim", 0.5, ""),
    (2, "Go-live readiness review", "Auto", 1, 22, "FS", None, 23, "FS", None, None, None, None, None, None, "Priya Sharma", None, ""),
    (2, "Go-live", "Auto", 0, 24, "FS", None, None, None, None, None, D(2026, 12, 18), None, None, None, None, None, "Milestone with a deadline of 18-Dec."),
    (2, "Hypercare support", "Auto", 10, 25, "FS", None, None, None, None, None, None, None, None, None, "Jordan Lee", 0.5, ""),
    (1, "Project Close", "Auto", None, None, None, None, None, None, None, None, None, None, None, None, None, None, ""),
    (2, "Steering committee update", "Manual", 1, None, None, None, None, None, None, D(2026, 11, 16), None, None, None, None, "Priya Sharma", None, "Manual mode example: fixed on the date in the Start No Earlier Than / Manual Start column."),
    (2, "Lessons learned", "Auto", 2, 26, "FS", None, None, None, None, None, None, None, None, None, "Priya Sharma", None, ""),
    (2, "Project closed", "Auto", 0, 29, "FS", None, None, None, None, None, None, None, None, None, None, None, "Milestone"),
]
IN_KEYS = ["Level", "Name", "Mode", "Dur", "P1", "T1", "L1", "P2", "T2", "L2", "Cons", "DL", "AS", "AF", "Pct", "Res", "Units", "Notes"]
DATE_KEYS = {"Cons", "DL", "PS", "PF", "AS", "AF", "FS", "FF", "SrcS", "SrcF", "BS", "BF"}
HELPER_DATES = {"SN", "FN", "D1P", "D2P", "D1F", "D2F", "CPS", "CPF"}
CENTER = {"ID", "UID", "WBS", "Level", "Mode", "Dur", "P1", "T1", "L1", "P2", "T2", "L2", "SVar", "Var", "DLChk",
          "Status", "Sprint", "SrcD", "BD", "DVar"}
for i in range(N_TASKS):
    r = FIRST + i
    f = tasks_formulas(r)
    for k, (key, h, kind, w) in zip(C, COLS):
        cell = ts[f"{C[k]}{r}"]
        cell.border = BOX
        if kind == "in":
            cell.font = INPUT_FONT
        else:
            cell.font = F(color="404040") if kind == "help" else F()
            cell.fill = CALC_FILL
            if k in f:
                cell.value = f[k]
        if k in DATE_KEYS:
            cell.number_format = DATE
        elif k in HELPER_DATES:
            cell.number_format = "dd-mmm-yy;;-"
        elif k in ("Pct", "Prog", "Units"):
            cell.number_format = "0%"
        elif k in ("Cost", "SrcC", "BC", "CVar"):
            cell.number_format = "#,##0.00"
        elif k in ("Work", "WPD", "SrcW", "BW", "WVar"):
            cell.number_format = "#,##0.0"
        if k in CENTER:
            cell.alignment = Alignment(horizontal="center")
    ts[f"{C['UID']}{r}"] = i + 1      # stable IDs: links keep working when rows move
    if i < len(SAMPLE):
        for k, v in zip(IN_KEYS, SAMPLE[i]):
            if v is not None:
                ts[f"{C[k]}{r}"] = v

ts.freeze_panes = f"{C['Mode']}{FIRST}"
ts.column_dimensions.group(C["Lvl"], C["MKey"], hidden=True, outline_level=1)
ts.auto_filter.ref = f"A5:{C['CVar']}{LAST}"


def col_rng(k):
    return f"{C[k]}{FIRST}:{C[k]}{LAST}"


def add_dv(ws, dv, *ranges):
    ws.add_data_validation(dv)
    for x in ranges:
        dv.add(x)


add_dv(ts, DataValidation(type="whole", operator="greaterThanOrEqual", formula1="1", showErrorMessage=True,
                          errorTitle="UID", error="UID must be a whole number; give each task a different one (see Settings > NextFreeUID)."), col_rng("UID"))
add_dv(ts, DataValidation(type="whole", operator="between", formula1="1", formula2="9", showErrorMessage=True,
                          errorTitle="Level", error="Outline level must be 1-9 and increase by at most 1 from the row above."), col_rng("Level"))
add_dv(ts, DataValidation(type="list", formula1="=L_Mode"), col_rng("Mode"))
add_dv(ts, DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", showErrorMessage=True,
                          error="Duration is in working days (0 = milestone)."), col_rng("Dur"))
add_dv(ts, DataValidation(type="whole", operator="greaterThanOrEqual", formula1="1", showErrorMessage=True,
                          error="Enter the UID of another task. Problems are explained in the Link Check column."), col_rng("P1"), col_rng("P2"))
add_dv(ts, DataValidation(type="list", formula1="=L_DepType"), col_rng("T1"), col_rng("T2"))
add_dv(ts, DataValidation(type="whole", operator="between", formula1="-365", formula2="365"), col_rng("L1"), col_rng("L2"))
add_dv(ts, DataValidation(type="date", operator="greaterThan", formula1="1"), col_rng("Cons"), col_rng("DL"), col_rng("AS"), col_rng("AF"))
add_dv(ts, DataValidation(type="decimal", operator="between", formula1="0", formula2="1", showErrorMessage=True,
                          error="Enter 0% - 100%."), col_rng("Pct"))
add_dv(ts, DataValidation(type="list", formula1="=res_Name"), col_rng("Res"))
add_dv(ts, DataValidation(type="decimal", operator="between", formula1="0", formula2="10"), col_rng("Units"))

# conditional formatting on Tasks
full = f"A{FIRST}:{C['CVar']}{LAST}"
ts.conditional_formatting.add(full, FormulaRule(formula=[f"${C['IsSum']}{FIRST}=TRUE"], font=Font(bold=True), fill=fill("DDEBF7")))
for k in ("Dur", "P1", "T1", "L1", "P2", "T2", "L2", "Cons", "AS", "AF", "Pct", "Res", "Units"):
    ts.conditional_formatting.add(col_rng(k), FormulaRule(formula=[f"${C['IsSum']}{FIRST}=TRUE"], font=Font(color="A6A6A6", italic=True, bold=True)))
STATUS_COL = {"Complete": "548235", "In Progress": "2F75B5", "Behind": "C00000", "Late Start": "C55A11", "Not Started": "7F7F7F"}
for s, colr in STATUS_COL.items():
    ts.conditional_formatting.add(col_rng("Status"), FormulaRule(formula=[f'${C["Status"]}{FIRST}="{s}"'], font=Font(color=colr, bold=True)))
for k in ("SVar", "Var", "DVar", "WVar", "CVar"):
    ts.conditional_formatting.add(col_rng(k), FormulaRule(formula=[f'AND(ISNUMBER(${C[k]}{FIRST}),${C[k]}{FIRST}>0)'], font=Font(color="C00000", bold=True)))
    ts.conditional_formatting.add(col_rng(k), FormulaRule(formula=[f'AND(ISNUMBER(${C[k]}{FIRST}),${C[k]}{FIRST}<0)'], font=Font(color="548235", bold=True)))
ts.conditional_formatting.add(col_rng("Name"), FormulaRule(formula=[f'AND(${C["IsLeaf"]}{FIRST}=1,N(${C["Dur"]}{FIRST})=0)'], font=Font(italic=True, color="7030A0")))
ts.conditional_formatting.add(col_rng("WBS"), FormulaRule(formula=[f'${C["WBS"]}{FIRST}="?"'], fill=fill("FFC7CE")))
ts.conditional_formatting.add(col_rng("UID"), FormulaRule(
    formula=[f'OR(AND(${C["Name"]}{FIRST}<>"",${C["UID"]}{FIRST}=""),AND(${C["UID"]}{FIRST}<>"",COUNTIF(tsk_UID,${C["UID"]}{FIRST})>1))'],
    fill=fill("FFC7CE"), font=Font(color="9C0006", bold=True)))
ts.conditional_formatting.add(col_rng("LinkChk"), FormulaRule(formula=[f'${C["LinkChk"]}{FIRST}<>""'], fill=fill("FFC7CE"), font=Font(color="9C0006")))
ts.conditional_formatting.add(col_rng("DLChk"), FormulaRule(formula=[f'LEFT(${C["DLChk"]}{FIRST},4)="Late"'], font=Font(color="C00000", bold=True)))
ts.conditional_formatting.add(col_rng("DLChk"), FormulaRule(formula=[f'${C["DLChk"]}{FIRST}="OK"'], font=Font(color="548235", bold=True)))
ts.sheet_view.zoomScale = 90

# ---------------------------------------------------------------- Sprints
sp = WB.create_sheet("Sprints")
sp["A1"] = '=ProjectName&"  -  Sprint & Capacity Planning"'
sp["A1"].font = F(size=16, bold=True, color="1F4E78")
sp["A2"] = ("Sprint calendar is generated from Settings (Sprint1Start, SprintLength, SprintCount). Work is prorated across sprints "
            "by the calendar days each task's forecast dates overlap the sprint.")
sp["A2"].font = F(italic=True, color="595959")
SCOLS = [("#", "calc", 5), ("Sprint", "calc", 12), ("Start", "calc", 14), ("End", "calc", 14), ("Working Days", "calc", 9),
         ("Sprint Goal", "in", 36), ("Tasks Starting", "calc", 9), ("Planned Work (hrs)", "calc", 11),
         ("Team Capacity (hrs)", "calc", 11), ("Utilisation", "calc", 10), ("Status", "calc", 12)]
for i, (t, k, w) in enumerate(SCOLS):
    header(sp, 5, 1 + i, t, k, w)
S0 = 6
GOALS = [] if BLANK else ["Charter & kickoff", "Requirements", "Design & architecture", "Foundations: DB + components", "Core APIs + screens",
         "Payments + screens", "Feature complete", "System testing", "UAT", "Release & hypercare", "Hypercare", "Close-out"]


def overlap_sum(sv, ev, extra=""):
    mn = f"(tsk_FinishN+{ev}-ABS(tsk_FinishN-{ev}))/2"
    mx = f"(tsk_StartN+{sv}+ABS(tsk_StartN-{sv}))/2"
    x = f"({mn}-{mx}+1)"
    return f"SUMPRODUCT({extra}tsk_WorkPerDay*(({x}+ABS{x})/2))"


for i in range(N_SPR):
    r = S0 + i
    n = i + 1
    vals = {
        1: n,
        2: f'=IF($A{r}<=SprintCount,SprintPrefix&$A{r},"")',
        3: f"=Sprint1Start+($A{r}-1)*SprintLength",
        4: f"=$C{r}+SprintLength-1",
        5: f"=NETWORKDAYS.INTL($C{r},$D{r},{WD})",
        6: GOALS[i] if i < len(GOALS) else None,
        7: f'=IF($B{r}="","",COUNTIF(tsk_Sprint,$B{r}))',
        8: f'=IF($B{r}="","",{overlap_sum(f"$C{r}", f"$D{r}")})',
        9: f'=IF($B{r}="","",SUMIFS(res_MaxUnits,res_Type,"Work")*HoursPerDay*$E{r})',
        10: f'=IF(OR($B{r}="",N($I{r})=0),"",$H{r}/$I{r})',
        11: f'=IF($B{r}="","",IF(StatusDate>$D{r},"Closed",IF(StatusDate>=$C{r},"Active","Future")))',
    }
    for c, v in vals.items():
        cell = sp.cell(r, c, v)
        cell.border = BOX
        cell.font = INPUT_FONT if c == 6 else F()
        if c != 6:
            cell.fill = CALC_FILL
    sp.cell(r, 3).number_format = DATE
    sp.cell(r, 4).number_format = DATE
    sp.cell(r, 8).number_format = "#,##0.0"
    sp.cell(r, 9).number_format = "#,##0.0"
    sp.cell(r, 10).number_format = "0%"
SL = S0 + N_SPR - 1
name("spr_Num", f"Sprints!$A${S0}:$A${SL}")
name("spr_Name", f"Sprints!$B${S0}:$B${SL}")
name("spr_Start", f"Sprints!$C${S0}:$C${SL}")
name("spr_End", f"Sprints!$D${S0}:$D${SL}")
name("spr_Capacity", f"Sprints!$I${S0}:$I${SL}")
name("spr_Work", f"Sprints!$H${S0}:$H${SL}")
sp.conditional_formatting.add(f"A{S0}:K{SL}", FormulaRule(formula=[f'$B{S0}=""'], font=Font(color="D9D9D9")))
sp.conditional_formatting.add(f"A{S0}:K{SL}", FormulaRule(formula=[f'$K{S0}="Active"'], fill=fill("FFF2CC"), font=Font(bold=True)))
sp.conditional_formatting.add(f"J{S0}:J{SL}", FormulaRule(formula=[f'AND(ISNUMBER($J{S0}),$J{S0}>1)'], font=Font(color="C00000", bold=True), fill=fill("FFC7CE")))

# resource x sprint matrix
M_HDR = SL + 3          # title row
M_NAME = M_HDR + 1      # sprint name row
M_START = M_HDR + 2
M_END = M_HDR + 3
M0 = M_HDR + 4
N_MRES = 25
sp.cell(M_HDR, 1, "Resource load by sprint (hours)  -  red = over capacity (Max Units x Hours/day x working days)").font = F(bold=True, size=12, color="1F4E78")
header(sp, M_NAME, 2, "Resource", "calc")
header(sp, M_NAME, 3, "Max Units", "calc")
sp.cell(M_START, 2, "Sprint start").font = F(italic=True)
sp.cell(M_END, 2, "Sprint end").font = F(italic=True)
U0 = M0 + N_MRES + 2  # utilisation matrix
sp.cell(U0 - 1, 1, "Resource utilisation by sprint (%)").font = F(bold=True, size=12, color="1F4E78")
header(sp, U0, 2, "Resource", "calc")
header(sp, U0, 3, "Max Units", "calc")
for j in range(N_SPR):
    col = 4 + j
    L = CL(col)
    header(sp, M_NAME, col, f"=INDEX(spr_Name,{j + 1})", "calc")
    header(sp, U0, col, f"={L}${M_NAME}", "calc")
    sp.cell(M_START, col, f"=INDEX(spr_Start,{j + 1})").number_format = "dd-mmm"
    sp.cell(M_END, col, f"=INDEX(spr_End,{j + 1})").number_format = "dd-mmm"
    if sp.column_dimensions[L].width is None or j >= 11:
        sp.column_dimensions[L].width = max(sp.column_dimensions[L].width or 0, 10)
    for i in range(N_MRES):
        r = M0 + i
        c = sp.cell(r, col, f'=IF(OR($B{r}="",{L}${M_NAME}=""),"",'
                            f'{overlap_sum(f"{L}${M_START}", f"{L}${M_END}", f"(tsk_Resource=$B{r})*")})')
        c.number_format = "#,##0;;-"
        c.border = BOX
        c.font = F()
        ru = U0 + 1 + i
        u = sp.cell(ru, col)
        u.number_format = "0%;;-"
        u.border = BOX
        u.font = F()
for i in range(N_MRES):
    r = M0 + i
    sp.cell(r, 2, f'=IF(INDEX(res_Name,{i + 1})="","",INDEX(res_Name,{i + 1}))').font = F(bold=True)
    sp.cell(r, 3, f'=IF($B{r}="","",INDEX(res_MaxUnits,{i + 1}))').number_format = "0%"
    ru = U0 + 1 + i
    sp.cell(ru, 2, f"=$B{r}").font = F(bold=True)
    sp.cell(ru, 3, f"=$C{r}").number_format = "0%"
last_col = CL(3 + N_SPR)
sp.conditional_formatting.add(
    f"D{M0}:{last_col}{M0 + N_MRES - 1}",
    FormulaRule(formula=[f'AND(ISNUMBER(D{M0}),D{M0}>$C{M0}*HoursPerDay*INDEX(Sprints!$E${S0}:$E${SL},COLUMN()-3))'],
                fill=fill("FFC7CE"), font=Font(color="9C0006", bold=True)))
sp.conditional_formatting.add(
    f"D{U0 + 1}:{last_col}{U0 + N_MRES}",
    ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF", mid_type="num", mid_value=0.8,
                   mid_color="C6EFCE", end_type="num", end_value=1.2, end_color="F8696B"))
sp.freeze_panes = "D6"
sp.column_dimensions["B"].width = 16

# fix the utilisation formula to the clean form (capacity = units x hrs/day x working days)
for j in range(N_SPR):
    col = 4 + j
    L = CL(col)
    for i in range(N_MRES):
        r = M0 + i
        ru = U0 + 1 + i
        sp.cell(ru, col).value = (f'=IF(OR($B{ru}="",{L}${M_NAME}="",N($C{ru})=0),"",'
                                  f'{L}{r}/($C{ru}*HoursPerDay*INDEX(Sprints!$E${S0}:$E${SL},{j + 1})))')

# ---------------------------------------------------------------- Gantt
gs = WB.create_sheet("Gantt", 2)
gs["A1"] = '=ProjectName&"  -  Gantt Chart"'
gs["A1"].font = F(size=16, bold=True, color="1F4E78")
gs["A2"] = '="Status date: "&TEXT(StatusDate,"ddd dd-mmm-yyyy")&"   |   Forecast finish: "&TEXT(MAX(tsk_FcFinish),"dd-mmm-yyyy")&"   |   Planned finish: "&TEXT(MAX(tsk_PlanFinish),"dd-mmm-yyyy")'
gs["A2"].font = F(bold=True, color="404040")
GT = [("ID", 5), ("WBS", 7), ("Task Name", 34), ("Dur", 5), ("Planned Start", 10), ("Planned Finish", 10),
      ("Fcst / Act Start", 10), ("Fcst / Act Finish", 10), ("%", 6), ("Resource", 13), ("Status", 11),
      ("IsSum", 5), ("IsMile", 5), ("Done From", 8), ("Done To", 8), ("Rem From", 8), ("Rem To", 8), ("Late", 5),
      ("Deadline", 8), ("", 1)]
for i, (t, w) in enumerate(GT):
    header(gs, 5, i + 1, t, "help" if 12 <= i + 1 <= 19 else "calc", w)
gs.row_dimensions[5].height = 30
G0 = len(GT) + 1  # first timeline column (T)
GL = CL(G0 + N_GANTT - 1)
gs.column_dimensions.group("L", "S", hidden=True, outline_level=1)
# legend
LEG = [("━", None, "1F3864", "Planned dates (line)"), ("", "70AD47", None, "Done (% complete)"),
       ("", "9DC3E6", None, "Remaining - on track"), ("", "F4B183", None, "Remaining - late vs plan"),
       ("", "404040", None, "Summary (forecast)"), ("◆", None, "1F3864", "Forecast milestone (red = late)"),
       ("◇", None, "1F3864", "Planned milestone"), ("▼", None, "C00000", "Deadline"),
       ("", None, None, "Red column = status date; grey = non-working")]

for n, (ch, colr, fc, txt) in enumerate(LEG):
    row, col = 1 + n // 5, G0 + (n % 5) * 11
    sw = gs.cell(row, col, ch or None)
    sw.alignment = Alignment(horizontal="center")
    sw.font = F(bold=True, color=fc or "1F3864")
    if colr:
        sw.fill = fill(colr)
    if n == 8:
        sw.border = Border(left=Side(style="medium", color="FF0000"))
    lab = gs.cell(row, col + 1, txt)
    lab.font = F(size=9)

for j in range(N_GANTT):
    col = G0 + j
    L = CL(col)
    prev = CL(col - 1)
    gs.column_dimensions[L].width = 2.9
    gs.cell(5, col, "=GanttStart" if j == 0 else f"={prev}5+GanttPeriod").number_format = "d"
    if j == 0:
        mon = f'=TEXT({L}5,"mmm")'
    elif j == 1:
        mon = f'=IF(MONTH({L}5)<>MONTH({prev}5),TEXT({L}5,"mmm"),"\'"&TEXT({prev}5,"yy"))'
    else:
        pp = CL(col - 2)
        mon = f'=IF(MONTH({L}5)<>MONTH({prev}5),TEXT({L}5,"mmm"),IF(MONTH({prev}5)<>MONTH({pp}5),"\'"&TEXT({prev}5,"yy"),""))'
    gs.cell(4, col, mon)
    gs.cell(3, col, f'=IF(GanttPeriod=1,LEFT(TEXT({L}5,"ddd"),1),"")')
    for rr in (3, 4, 5):
        c = gs.cell(rr, col)
        c.font = F(size=8, bold=(rr == 4), color="FFFFFF" if rr == 5 else "404040")
        c.alignment = Alignment(horizontal="center" if rr != 4 else "left", vertical="center", text_rotation=0)
        if rr == 5:
            c.fill = HDR_CALC

def task_refs(n):
    """References to the n-th task row by named range, so inserted Tasks rows still line up."""
    return {k: f"INDEX({nm},{n})" for k, nm in TSK_NAMES.items()}


for i in range(N_TASKS):
    r = FIRST + i
    t = task_refs(i + 1)
    blank = f'{t["Name"]}=""'
    vals = [
        f'=IF({blank},"",{t["ID"]})', f'=IF({blank},"",{t["WBS"]})',
        f'=IF({blank},"",REPT("    ",{t["Lvl"]}-1)&{t["Name"]})',
        f'=IF({blank},"",IF({t["IsSum"]},NETWORKDAYS.INTL({t["FS"]},{t["FF"]},{WD}),N({t["Dur"]})))',
        f'=IF({blank},"",{t["PS"]})', f'=IF({blank},"",{t["PF"]})',
        f'=IF({blank},"",{t["FS"]})', f'=IF({blank},"",{t["FF"]})',
        f'=IF({blank},"",{t["Prog"]})', f'=IF({blank},"",{t["Res"]}&"")', f'=IF({blank},"",{t["Status"]})',
        f'=IF({blank},FALSE(),{t["IsSum"]})',
        f'=AND({t["IsLeaf"]}=1,N({t["Dur"]})=0)',
        f'=IF(AND({t["IsLeaf"]}=1,NOT($M{r}),N($I{r})>0),$G{r},0)',
        f'=IF($N{r}=0,0,IF($I{r}>=1,$H{r},MAX($N{r},$G{r}+ROUND(($H{r}-$G{r}+1)*$I{r},0)-1)))',
        f'=IF(OR({t["IsLeaf"]}<>1,$M{r},N($I{r})>=1),0,IF($O{r}>0,$O{r}+1,$G{r}))',
        f'=IF($P{r}=0,0,$H{r})',
        f'=AND({t["Name"]}<>"",N({t["Var"]})>0)',
        f'=IF({blank},0,N({t["DL"]}))',
    ]
    for c, v in enumerate(vals, start=1):
        cell = gs.cell(r, c, v)
        cell.font = F(size=9)
        cell.border = BOX
        if c in (5, 6, 7, 8):
            cell.number_format = "dd-mmm-yy"
        elif c in (14, 15, 16, 17, 19):
            cell.number_format = "dd-mmm;;-"
        elif c == 9:
            cell.number_format = "0%"
        if c in (1, 2, 4, 9, 11):
            cell.alignment = Alignment(horizontal="center")
    for j in range(N_GANTT):
        L = CL(G0 + j)
        hd = f"{L}$5"
        in_plan = f"AND({hd}<=$F{r},{hd}+GanttPeriod-1>=$E{r})"
        in_p = lambda d: f"AND({d}>={hd},{d}<{hd}+GanttPeriod)"
        dl = f'AND($S{r}>0,{in_p(f"$S{r}")})'
        glyph = (f'IF($M{r},IF({in_p(f"$G{r}")},"◆",IF({in_p(f"$E{r}")},"◇",IF({dl},"▼",""))),'
                 f'IF({dl},"▼",IF({in_plan},"━","")))')
        gs.cell(r, G0 + j, f'=IF($E{r}="","",{glyph})')
grid = f"{CL(G0)}{FIRST}:{GL}{LAST}"
g1 = CL(G0)
ov = lambda a, b: f"AND({a}>0,{g1}$5<={b},{g1}$5+GanttPeriod-1>={a})"
rules = [
    FormulaRule(formula=[f"AND($L{FIRST},{g1}$5<=$H{FIRST},{g1}$5+GanttPeriod-1>=$G{FIRST})"], fill=fill("404040"), font=Font(color="FFFFFF")),
    FormulaRule(formula=[ov(f"$N{FIRST}", f"$O{FIRST}")], fill=fill("70AD47")),
    FormulaRule(formula=[f"AND($R{FIRST},{ov(f'$P{FIRST}', f'$Q{FIRST}')})"], fill=fill("F4B183")),
    FormulaRule(formula=[ov(f"$P{FIRST}", f"$Q{FIRST}")], fill=fill("9DC3E6")),
    FormulaRule(formula=[f"AND($M{FIRST},$R{FIRST})"], font=Font(color="C00000", bold=True)),
    FormulaRule(formula=[f"AND($S{FIRST}>0,$S{FIRST}>={g1}$5,$S{FIRST}<{g1}$5+GanttPeriod)"], font=Font(color="C00000", bold=True)),
    FormulaRule(formula=[f"AND(StatusDate>={g1}$5,StatusDate<{g1}$5+GanttPeriod)"],
                border=Border(left=Side(style="thin", color="FF0000"), right=Side(style="thin", color="FF0000"))),
    FormulaRule(formula=[f"AND(GanttPeriod=1,NETWORKDAYS.INTL({g1}$5,{g1}$5,{WD})=0)"], fill=fill("EDEDED")),
]
for rule in rules:
    gs.conditional_formatting.add(grid, rule)
for rg in (f"{CL(G0)}5:{GL}5",):
    gs.conditional_formatting.add(rg, FormulaRule(formula=[f"AND(StatusDate>={g1}$5,StatusDate<{g1}$5+GanttPeriod)"], fill=fill("C00000")))
gs.conditional_formatting.add(f"A{FIRST}:K{LAST}", FormulaRule(formula=[f"$L{FIRST}"], font=Font(bold=True), fill=fill("DDEBF7")))
for s, colr in STATUS_COL.items():
    gs.conditional_formatting.add(f"K{FIRST}:K{LAST}", FormulaRule(formula=[f'$K{FIRST}="{s}"'], font=Font(color=colr, bold=True)))
for r in range(FIRST, LAST + 1):
    for j in range(N_GANTT):
        c = gs.cell(r, G0 + j)
        c.font = Font(name=FONT, size=9, bold=True, color="1F3864")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = Border(left=Side(style="hair", color="F2F2F2"), right=Side(style="hair", color="F2F2F2"))
gs.freeze_panes = f"{CL(G0)}{FIRST}"
gs.sheet_view.zoomScale = 90

# ---------------------------------------------------------------- MSP export
mx = WB.create_sheet("MSP_Tasks")
MSP = [("ID", "ID", 5), ("Unique ID", None, 8), ("Name", None, 36), ("Outline Level", "Lvl", 8), ("Task Mode", None, 18),
       ("Duration", None, 9), ("Start", "FS", 12), ("Finish", "FF", 12), ("Predecessors", "PredTxt", 14),
       ("Unique ID Predecessors", None, 14), ("Resource Names", None, 22), ("% Complete", None, 9),
       ("Actual Start", None, 12), ("Actual Finish", None, 12), ("Baseline Start", "PS", 12), ("Baseline Finish", "PF", 12),
       ("Baseline Duration", None, 9), ("Baseline Work", None, 9), ("Baseline Cost", None, 11),
       ("Constraint Type", None, 22), ("Constraint Date", None, 12), ("Deadline", None, 12), ("Notes", "Notes", 40)]
for i, (h, k, w) in enumerate(MSP):
    header(mx, 1, i + 1, h, "calc", w)
for i in range(N_TASKS):
    r = 2 + i
    t = task_refs(i + 1)
    bl = f'{t["Name"]}=""'
    row = {
        "ID": f'=IF({bl},"",{t["ID"]})',
        "Unique ID": f'=IF({bl},"",{t["UID"]})',
        "Unique ID Predecessors": f'=IF(OR({bl},{t["IsSum"]}),"",{t["UPredTxt"]})',
        "Baseline Duration": f'=IF(OR({bl},{t["BD"]}=""),"",N({t["BD"]})&"d")',
        "Baseline Work": f'=IF(OR({bl},{t["BW"]}=""),"",N({t["BW"]})&"h")',
        "Baseline Cost": f'=IF(OR({bl},{t["BC"]}=""),"",N({t["BC"]}))',
        "Deadline": f'=IF(OR({bl},N({t["DL"]})=0),"",{t["DL"]})',
        "Name": f'=IF({bl},"",{t["Name"]})',
        "Outline Level": f'=IF({bl},"",{t["Lvl"]})',
        "Task Mode": f'=IF({bl},"",IF({t["Mode"]}="Manual","Manually Scheduled","Auto Scheduled"))',
        "Duration": f'=IF(OR({bl},{t["IsSum"]}),"",N({t["Dur"]})&"d")',
        "Start": f'=IF({bl},"",{t["FS"]})',
        "Finish": f'=IF({bl},"",{t["FF"]})',
        "Predecessors": f'=IF(OR({bl},{t["IsSum"]}),"",{t["PredTxt"]})',
        "Resource Names": f'=IF(OR({bl},{t["Res"]}=""),"",{t["Res"]}&IF(OR({t["Units"]}="",{t["Units"]}=1),"","["&ROUND({t["Units"]}*100,0)&"%]"))',
        "% Complete": f'=IF(OR({bl},{t["IsSum"]}),"",ROUND({t["Prog"]}*100,0))',
        "Actual Start": f'=IF(OR({bl},N({t["AS"]})=0),"",{t["AS"]})',
        "Actual Finish": f'=IF(OR({bl},N({t["AF"]})=0),"",{t["AF"]})',
        "Baseline Start": f'=IF({bl},"",{t["PS"]})',
        "Baseline Finish": f'=IF({bl},"",{t["PF"]})',
        "Constraint Type": f'=IF(OR({bl},{t["IsSum"]}),"",IF(AND({t["Mode"]}<>"Manual",N({t["Cons"]})>0),"Start No Earlier Than","As Soon As Possible"))',
        "Constraint Date": f'=IF(OR({bl},{t["IsSum"]},N({t["Cons"]})=0,{t["Mode"]}="Manual"),"",{t["Cons"]})',
        "Notes": f'=IF({bl},"",{t["Notes"]}&"")',
    }
    for c, (h, k, w) in enumerate(MSP, start=1):
        cell = mx.cell(r, c, row[h])
        cell.font = F()
        if h in ("Start", "Finish", "Actual Start", "Actual Finish", "Baseline Start", "Baseline Finish", "Constraint Date", "Deadline"):
            cell.number_format = "dd-mmm-yyyy"
mx.freeze_panes = "C2"
name("msp_Tasks", f"MSP_Tasks!$A$1:${CL(len(MSP))}${N_TASKS + 1}")

mr = WB.create_sheet("MSP_Resources")
MR = [("Name", "res_Name", 22), ("Initials", "res_Initials", 9), ("Group", "res_Group", 20), ("Type", "res_Type", 10),
      ("Max Units", "res_MaxUnits", 10), ("Standard Rate", "res_Rate", 12), ("Email Address", "res_Email", 26)]
for i, (h, n, w) in enumerate(MR):
    header(mr, 1, i + 1, h, "calc", w)
for i in range(N_RES):
    r = 2 + i
    for c, (h, n, w) in enumerate(MR, start=1):
        if h == "Max Units":
            v = f'=IF(INDEX(res_Name,{i + 1})="","",ROUND(INDEX({n},{i + 1})*100,0)&"%")'
        elif h == "Standard Rate":
            v = f'=IF(INDEX(res_Name,{i + 1})="","",INDEX({n},{i + 1})&"/h")'
        else:
            v = f'=IF(INDEX({n},{i + 1})="","",INDEX({n},{i + 1})&"")'
        mr.cell(r, c, v).font = F()

# ---------------------------------------------------------------- Dashboard
db = WB.create_sheet("Dashboard", 0)
db["A1"] = '=ProjectName&"  -  Dashboard"'
db["A1"].font = F(size=18, bold=True, color="1F4E78")
db["A2"] = '="Project manager: "&ProjectManager&"     Status date: "&TEXT(StatusDate,"ddd dd-mmm-yyyy")'
db["A2"].font = F(italic=True, color="595959")
db.column_dimensions["A"].width = 3
db.column_dimensions["B"].width = 38
db.column_dimensions["C"].width = 18
db.column_dimensions["D"].width = 4
db.column_dimensions["E"].width = 18
db.column_dimensions["F"].width = 12
KPI = [
    ("Planned start", "=MIN(tsk_PlanStart)", DATE),
    ("Planned finish", "=MAX(tsk_PlanFinish)", DATE),
    ("Forecast finish", "=MAX(tsk_FcFinish)", DATE),
    ("Finish variance (work days)", f"=IF(MAX(tsk_FcFinish)>=MAX(tsk_PlanFinish),NETWORKDAYS.INTL(MAX(tsk_PlanFinish),MAX(tsk_FcFinish),{WD})-1,1-NETWORKDAYS.INTL(MAX(tsk_FcFinish),MAX(tsk_PlanFinish),{WD}))", "+0;-0;0"),
    ("Overall progress (duration-weighted)", "=IFERROR(SUM(tsk_WtDone)/SUM(tsk_WtDur),0)", "0%"),
    ("Planned progress by status date", f"=IFERROR(SUMPRODUCT(tsk_WtDur*(tsk_PlanFinish<StatusDate))/SUM(tsk_WtDur),0)", "0%"),
    ("Detail tasks", "=SUM(tsk_IsLeaf)", "0"),
    ("Milestones", '=SUMPRODUCT(tsk_IsLeaf*(tsk_Duration=0)*(tsk_Name<>""))', "0"),
    ("Total work (hrs)", "=SUMPRODUCT(tsk_Work,tsk_IsLeaf)", "#,##0"),
    ("Total cost", "=SUMPRODUCT(tsk_Cost,tsk_IsLeaf)", "#,##0.00"),
    ("Current sprint", '=IF(OR(StatusDate<Sprint1Start,StatusDate>SprintLastEnd),"(outside sprint range)",INDEX(spr_Name,MATCH(StatusDate,spr_Start,1)))', "@"),
    ("Baseline", '=IF(COUNT(tsk_BaseStart)=0,"Not set (planned = calculated)","Set on "&COUNT(tsk_BaseStart)&" tasks")', "@"),
    ("Missed deadlines", '=COUNTIF(tsk_DeadlineCheck,"Late*")', "0"),
    ("Link problems (see Tasks > Link Check)", '=COUNTIF(tsk_LinkCheck,"?*")', "0"),
]
header(db, 4, 2, "Key indicator", "calc")
header(db, 4, 3, "Value", "calc")
for i, (lbl, f, fmt) in enumerate(KPI):
    r = 5 + i
    db.cell(r, 2, lbl).font = F(bold=True)
    c = db.cell(r, 3, f)
    c.number_format = fmt
    c.font = F(size=11, bold=True)
    c.alignment = Alignment(horizontal="right")
    for cc in (2, 3):
        db.cell(r, cc).border = BOX
db.conditional_formatting.add("C8", FormulaRule(formula=["C8>0"], font=Font(color="C00000", bold=True)))
db.conditional_formatting.add("C8", FormulaRule(formula=["C8<=0"], font=Font(color="548235", bold=True)))
for rr in (17, 18):
    db.conditional_formatting.add(f"C{rr}", FormulaRule(formula=[f"C{rr}>0"], font=Font(color="C00000", bold=True), fill=fill("FFC7CE")))
header(db, 4, 5, "Status (detail tasks)", "calc")
header(db, 4, 6, "Count", "calc")
for i, s in enumerate(STATUS_COL):
    r = 5 + i
    db.cell(r, 5, s).font = F(bold=True, color=STATUS_COL[s])
    db.cell(r, 6, f'=COUNTIFS(tsk_Status,E{r},tsk_IsLeaf,1)').font = F(bold=True)
    for cc in (5, 6):
        db.cell(r, cc).border = BOX
ch = BarChart()
ch.type = "bar"
ch.style = 10
ch.title = "Detail tasks by status"
ch.legend = None
ch.add_data(Reference(db, min_col=6, min_row=4, max_row=9), titles_from_data=True)
ch.set_categories(Reference(db, min_col=5, min_row=5, max_row=9))
ch.height, ch.width = 7, 14
db.add_chart(ch, "E12")
ch2 = BarChart()
ch2.title = "Sprint planned work vs capacity (hrs)"
ch2.add_data(Reference(sp, min_col=8, min_row=5, max_row=5 + 12), titles_from_data=True)
ch2.add_data(Reference(sp, min_col=9, min_row=5, max_row=5 + 12), titles_from_data=True)
ch2.set_categories(Reference(sp, min_col=2, min_row=6, max_row=5 + 12))
ch2.height, ch2.width = 8, 22
db.add_chart(ch2, "B29")
# upcoming milestones
header(db, 21, 2, "Next milestones (forecast)", "calc")
header(db, 21, 3, "Forecast date", "calc")
for k in range(1, 6):
    r = 21 + k
    nth = f"SMALL(tsk_MilestoneKey,{k})"
    db.cell(r, 2, f'=IFERROR(INDEX(tsk_Name,MATCH({nth},tsk_MilestoneKey,0)),"")')
    db.cell(r, 3, f'=IFERROR(INDEX(tsk_FcFinish,MATCH({nth},tsk_MilestoneKey,0)),"")').number_format = DATE
    for cc in (2, 3):
        db.cell(r, cc).border = BOX
        db.cell(r, cc).font = F()

# ---------------------------------------------------------------- Instructions
ins = WB.create_sheet("Instructions", 0)
ins.column_dimensions["A"].width = 3
ins.column_dimensions["B"].width = 28
ins.column_dimensions["C"].width = 110
ins["B1"] = "Project Planner - MS Project-compatible Excel template"
ins["B1"].font = F(size=18, bold=True, color="1F4E78")
TXT = [
    ("HOW IT WORKS", None),
    ("1. Settings", "Set project start, working week (Mon-Sun Yes/No), hours/day, status date, Gantt scale and sprint calendar. Add holidays on 'Holidays'."),
    ("2. Resources", "List people/roles with Max Units and hourly rate. They appear in the Resource dropdown on Tasks and in sprint capacity."),
    ("3. Tasks", "One row per task. Enter Level (1 = phase/summary, 2 = task, 3 = sub-task...), Task Name, Duration in working days (0 = milestone), "
                 "and up to two predecessors by UID with link type (FS, SS, FF, SF) and lag in days (negative = lead). A row becomes a SUMMARY automatically "
                 "when the rows beneath it have a higher level; its dates, work, cost and progress roll up. Optional Deadline: Deadline Check shows 'Late by Nd'."),
    ("4. Set the baseline", "When the plan is agreed: on Tasks, select the 5 'Baseline source' columns (rows 6 to the last task), Copy, then select the first "
                            "'Baseline Start' cell and Paste Special > Values. Planned Start/Finish now stay frozen (like MS Project's Baseline) and the "
                            "variance columns compare against it. To re-baseline, repeat; to clear it, delete the 5 Baseline columns' contents."),
    ("5. Track progress", "Enter Actual Start, Actual Finish and % Complete. Forecast Start/Finish then use actuals and push every dependent task. "
                          "Start / Finish / Duration / Work / Cost Variance and Status flag slippage."),
    ("6. Gantt", "Read-only view: grey line = planned dates, green = done, blue = remaining on track, orange = remaining but late vs plan, dark = summary, "
                 "red border = status date. Change GanttStart / GanttPeriod (1 = days, 7 = weeks) on Settings to scroll or zoom."),
    ("7. Sprints", "Sprint calendar from Settings. Each task is tagged with the sprint its forecast start falls in. Work hours are prorated across sprints "
                   "per resource, compared against capacity (Max Units x Hours/day x working days); over-allocations turn red."),
    ("8. Dashboard", "KPIs (including baseline, missed deadlines, link problems), status counts, upcoming milestones and sprint load vs capacity."),
    ("CELL COLOURS", None),
    ("Blue text", "Input - type here."),
    ("Black text on grey", "Formula - do not overwrite. Summary rows grey-out inputs that are ignored."),
    ("Dark blue header", "Input column.  Dark grey header = calculated.  Light grey header = helper column (hidden, grouped)."),
    ("RULES (same as MS Project)", None),
    ("ID and UID", "ID = position of the task (renumbers when rows move, like MS Project). UID = permanent number that predecessors refer to, so links "
                   "survive inserting, deleting and moving rows. Every row comes with a UID; a duplicate or missing UID turns red."),
    ("Inserting tasks", "Select an empty task row near the bottom, Copy, select the row where the new task goes, right-click > Insert Copied Cells. "
                        "Type the new task over the blue cells and set its UID to Settings > NextFreeUID (also shown at the top of Tasks). "
                        "Inserting a plain empty row instead leaves it without formulas."),
    ("Outline levels", "Each row's Level may be at most 1 higher than the row above. A '?' in WBS (red) means a level jump."),
    ("Predecessors", "Enter the predecessor's UID. Links to the task itself, to its own summary, to an empty row or to a missing UID are ignored and "
                     "explained in the Link Check column (red). Summary rows ignore Duration / Predecessors / Actuals - they roll up. "
                     "A loop through several tasks (A > B > A) still causes a circular reference."),
    ("Mode", "Auto = scheduled from predecessors, project start and 'Start No Earlier Than'. Manual = pinned to the date in 'Start No Earlier Than / Manual Start'."),
    ("Rescheduling", "Settings > RescheduleAfterStatus = Yes moves incomplete work to after the status date (like MS Project 'Update Project'), so late tasks push successors."),
    ("Capacity", "Rows provided: 200 tasks, 50 resources, 26 sprints, 100 holidays. To extend, copy the last formula row down and extend the named ranges in Name Manager."),
    ("MOVING TO MICROSOFT PROJECT", None),
    ("One-time setting", "MS Project > File > Options > Trust Center > Trust Center Settings > Legacy Formats > 'Allow loading files with legacy or non-default file formats'."),
    ("Import", "In MS Project: File > Open > choose this .xlsx > Import Wizard > New map > 'As a new project' > tick Tasks (and Resources) > "
               "Tasks source worksheet = MSP_Tasks, Resources = MSP_Resources. Column headers already use MS Project field names, so the wizard maps them automatically. "
               "Map 'Predecessors' (by ID) for a new project; 'Unique ID' and 'Unique ID Predecessors' are for merging into an existing one."),
    ("Tip: Start/Finish", "For auto-scheduled tasks map Duration + Predecessors (+ Constraint Type/Date) and set Start/Finish to '(not mapped)' - importing Start/Finish "
                          "makes MS Project add fixed 'Start No Earlier Than' constraints. Map Baseline Start/Finish to keep the plan, and Actual Start/Finish/% Complete for progress."),
    ("Rows", "MSP_Tasks contains 200 rows; blank rows import as blank tasks - delete them in MS Project or copy only the filled rows to a new sheet before importing."),
    ("From MS Project to Excel", "MS Project > File > Save As > Excel Workbook > use the same field names, then paste Name / Outline Level / Duration / Predecessors etc. into the Tasks input columns "
                                 "(split '3FS+2d' into Pred ID, Type, Lag)."),
    ("NAMED RANGES (Formulas > Name Manager)", None),
    ("Settings", "ProjectStart, StatusDate, StatusWorkday, WorkWeek, Holidays, HoursPerDay, DefaultUnits, RescheduleAfterStatus, GanttStart, GanttPeriod, Sprint1Start, SprintLength, SprintCount, SprintPrefix, SprintLastEnd"),
    ("Task columns", "tsk_ID, tsk_UID, tsk_WBS, tsk_Level, tsk_Name, tsk_Duration, tsk_Deadline, tsk_PlanStart, tsk_PlanFinish, tsk_ActStart, tsk_ActFinish, "
                     "tsk_FcStart, tsk_FcFinish, tsk_StartVariance, tsk_Variance, tsk_DeadlineCheck, tsk_LinkCheck, tsk_Progress, tsk_Status, tsk_Resource, tsk_Work, "
                     "tsk_Cost, tsk_Sprint, tsk_BaseStart ... tsk_BaseCost, tsk_DurationVariance, tsk_WorkVariance, tsk_CostVariance, tsk_CalcStart, tsk_CalcFinish "
                     "... and hdr_* anchors used by OFFSET look-ups. Gantt and MSP_Tasks read Tasks through INDEX(tsk_..., n)."),
    ("Resources / Sprints", "res_Name, res_MaxUnits, res_Rate, ...  spr_Name, spr_Start, spr_End, spr_Capacity, spr_Work"),
    ("Key formulas", "Calculated Start = WORKDAY.INTL(MAX(ProjectStart, StartNoEarlierThan, predecessor drivers) - 1, 1, WorkWeek, Holidays).  "
                     "Pred driver (FS) = WORKDAY.INTL(OFFSET(hdr_CalcFinish, MATCH(PredUID, tsk_UID, 0), 0), 1 + Lag, WorkWeek, Holidays).  "
                     "Planned Start = Baseline Start if set, else Calculated Start.  "
                     "Summary Start = MIN of its descendant block (OFFSET from the row, height = Descendant Count via COUNTIF(tsk_WBS, WBS & \".*\"))."),
]
r = 3
for k, v in TXT:
    if v is None:
        r += 1
        c = ins.cell(r, 2, k)
        c.font = WHITE_B
        c.fill = HDR_IN
        ins.cell(r, 3).fill = HDR_IN
    else:
        ins.cell(r, 2, k).font = F(bold=True)
        c = ins.cell(r, 3, v)
        c.font = F()
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ins.cell(r, 2).alignment = Alignment(vertical="top")
        if k == "Blue text":
            c.font = INPUT_FONT
        if k == "Black text on grey":
            c.fill = CALC_FILL
    r += 1

for ws in WB.worksheets:
    ws.sheet_properties.tabColor = {"Instructions": "1F4E78", "Dashboard": "548235", "Settings": "BF8F00", "Tasks": "2F75B5",
                                    "Gantt": "2F75B5", "Resources": "7030A0", "Sprints": "7030A0", "Holidays": "BF8F00",
                                    "MSP_Tasks": "595959", "MSP_Resources": "595959", "Lists": "A6A6A6"}.get(ws.title, "A6A6A6")
order = ["Instructions", "Dashboard", "Settings", "Tasks", "Gantt", "Resources", "Sprints", "Holidays", "MSP_Tasks", "MSP_Resources", "Lists"]
WB._sheets = [WB[n] for n in order]
WB.active = 3
WB.calculation.fullCalcOnLoad = True
if ARGS.recalc:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise SystemExit("--recalc needs LibreOffice (soffice) on PATH")
    with tempfile.TemporaryDirectory() as tmp:
        raw = Path(tmp) / Path(OUT).name
        WB.save(raw)
        subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(Path(tmp) / "out"), str(raw)],
                       check=True, capture_output=True, timeout=900)
        shutil.copyfile(Path(tmp) / "out" / raw.name, OUT)
else:
    WB.save(OUT)
print("saved", OUT)
