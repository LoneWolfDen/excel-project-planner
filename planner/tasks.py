"""Tasks sheet: the Excel Table "Tasks", where all scheduling happens.

Every formula below is a calculated column: one formula per column, written with structured references
(R("Duration") is this row's Duration, shown by Excel as [@Duration]). Rows inserted inside the table get the
formulas automatically.

Reading another row's CALCULATED value (a predecessor's finish, the parent's WBS) uses OFFSET from the column's
header cell. Excel judges circular references by the ranges a formula names, and OFFSET names only the header
cell; it is also the only construct that stays correct when rows are inserted directly above (a plain range
ending at the row above would not grow). OFFSET is volatile, so those few columns recalculate after every edit;
docs/ARCHITECTURE.md records the measurements and why the alternatives were rejected.

Everything else avoids volatile functions:
  * detail-task values live in _Leaf... helper columns; summaries roll them up over the rows below them with
    INDEX(col, first):INDEX(col, last), and the leaf columns never read the summary columns;
  * look-ups of INPUT columns (UID, Task Name, ID) use INDEX/MATCH over the whole column.
A predecessor must be a detail task listed above its successor (the critical path's backward pass relies on it).
"""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font
from openpyxl.worksheet.datavalidation import DataValidation

from .core import (DATE, LABEL_FONT, NOTE_FONT, Col, Tab, add_table, col_range, fill, name, title)
from .sample import COLUMNS as SAMPLE_COLUMNS

TABLE = "Tasks"
HEADER_ROW = 5
T = Tab(TABLE)
WD = "WorkWeek,Holidays"

# key -> column name (= header text = structured-reference name)
N = {
    "ID": "ID", "UID": "UID", "WBS": "WBS", "Level": "Level", "Name": "Task Name", "Mode": "Mode", "Dur": "Duration",
    "P1": "Pred1 UID", "T1": "Pred1 Type", "L1": "Pred1 Lag", "P2": "Pred2 UID", "T2": "Pred2 Type", "L2": "Pred2 Lag",
    "Cons": "Start No Earlier Than", "DL": "Deadline",
    "PS": "Planned Start", "PF": "Planned Finish", "AS": "Actual Start", "AF": "Actual Finish",
    "FS": "Forecast Start", "FF": "Forecast Finish", "SVar": "Start Var", "Var": "Finish Var",
    "LS": "Late Start", "LF": "Late Finish", "TS": "Total Slack", "FSl": "Free Slack", "Crit": "Critical",
    "DLChk": "Deadline Check", "LinkChk": "Link Check", "Pct": "% Complete", "Prog": "Progress", "Status": "Status",
    "Res": "Resource", "Units": "Units", "OverAlloc": "Over-allocated", "Work": "Work", "Cost": "Cost", "ACin": "Actual Cost", "SprIn": "Planned Sprint",
    "Sprint": "Sprint", "SP": "Story Points",
    "PredTxt": "Predecessors", "Notes": "Notes",
    "SrcS": "BL Source Start", "SrcF": "BL Source Finish", "SrcD": "BL Source Duration", "SrcW": "BL Source Work",
    "SrcC": "BL Source Cost", "BS": "Baseline Start", "BF": "Baseline Finish", "BD": "Baseline Duration",
    "BW": "Baseline Work", "BC": "Baseline Cost", "DVar": "Duration Var", "WVar": "Work Var", "CVar": "Cost Var",
    "Budget": "Budget", "PV": "PV", "EV": "EV", "ACt": "AC",
    # helper columns (hidden)
    "Lvl": "_Level", "Idx": "_Row", "Parent": "_ParentRow", "NextLvl": "_NextLevel", "IsSum": "_IsSummary",
    "Desc": "_Descendants", "IsLeaf": "_IsLeaf", "R1": "_Pred1Pos", "R2": "_Pred2Pos",
    "M1": "_Pred1Problem", "M2": "_Pred2Problem",
    "D1P": "_Pred1PlanDriver", "D2P": "_Pred2PlanDriver", "D1F": "_Pred1FcstDriver", "D2F": "_Pred2FcstDriver",
    "LCS": "_LeafCalcStart", "LCF": "_LeafCalcFinish", "LFS": "_LeafFcstStart", "LFF": "_LeafFcstFinish",
    "LProg": "_LeafProgress", "LWork": "_LeafWork", "LCost": "_LeafCost",
    "F1": "_Pred1FinishKey", "F2": "_Pred2FinishKey", "K1": "_Pred1SlackKey", "K2": "_Pred2SlackKey",
    "SuccLF": "_SuccessorFinishLimit", "LastSucc": "_LastSuccessorRow",
    "LLF": "_LeafLateFinish", "LLS": "_LeafLateStart", "LTS": "_LeafTotalSlack", "LCrit": "_LeafCritical",
    "CPS": "_CalcStart", "CPF": "_CalcFinish",
    "SWD": "_StartWorkday", "FWD": "_FinishWorkday", "WPD": "_WorkPerWorkday", "WBefore": "_WorkBeforeStart",
    "WtDur": "_WeightDur", "WtDone": "_WeightDone",
    "T1x": "_Pred1IDText", "T2x": "_Pred2IDText", "U1x": "_Pred1UIDText", "U2x": "_Pred2UIDText",
    "UPredTxt": "_UIDPredecessors", "MKey": "_MilestoneKey",
    # earned value and the report lists
    "PlanPct": "_PlannedPct", "LBud": "_LeafBudget", "LPV": "_LeafPV", "LEV": "_LeafEV", "LAC": "_LeafAC",
    "Phase": "_Phase", "PhaseKey": "_PhaseKey", "MsKey": "_MilestoneOrder", "LateKey": "_LateKey", "StartKey": "_StartSoonKey",
    "DueKey": "_DueSoonKey", "CritKey": "_CriticalKey",
    # resource and agile reports
    "LWDone": "_LeafWorkDone", "DoneDate": "_DoneDate", "LPts": "_LeafPoints", "CurKey": "_CurrentSprintKey",
}

# defined names used by every other sheet (tsk_Duration = Tasks[Duration], ...)
NAMES = {
    "ID": "tsk_ID", "UID": "tsk_UID", "WBS": "tsk_WBS", "Lvl": "tsk_Level", "Name": "tsk_Name", "Mode": "tsk_Mode",
    "Dur": "tsk_Duration", "Cons": "tsk_Constraint", "DL": "tsk_Deadline", "PS": "tsk_PlanStart",
    "PF": "tsk_PlanFinish", "AS": "tsk_ActStart", "AF": "tsk_ActFinish", "FS": "tsk_FcStart", "FF": "tsk_FcFinish",
    "SVar": "tsk_StartVariance", "Var": "tsk_Variance", "LS": "tsk_LateStart", "LF": "tsk_LateFinish",
    "TS": "tsk_TotalSlack", "FSl": "tsk_FreeSlack", "Crit": "tsk_Critical", "DLChk": "tsk_DeadlineCheck", "LinkChk": "tsk_LinkCheck",
    "Pct": "tsk_PctInput", "Prog": "tsk_Progress", "Status": "tsk_Status", "Res": "tsk_Resource",
    "Units": "tsk_Units", "Work": "tsk_Work", "Cost": "tsk_Cost", "Sprint": "tsk_Sprint",
    "PredTxt": "tsk_Predecessors", "Notes": "tsk_Notes", "BS": "tsk_BaseStart", "BF": "tsk_BaseFinish",
    "BD": "tsk_BaseDuration", "BW": "tsk_BaseWork", "BC": "tsk_BaseCost", "SrcD": "tsk_SourceDuration",
    "DVar": "tsk_DurationVariance", "WVar": "tsk_WorkVariance", "CVar": "tsk_CostVariance",
    "IsSum": "tsk_IsSummary", "IsLeaf": "tsk_IsLeaf", "CPS": "tsk_CalcStart", "CPF": "tsk_CalcFinish",
    "LWork": "tsk_LeafWork", "SWD": "tsk_StartWD", "FWD": "tsk_FinishWD", "WPD": "tsk_WorkPerWD",
    "WBefore": "tsk_WorkBeforeStart", "WtDur": "tsk_WtDur", "WtDone": "tsk_WtDone",
    "UPredTxt": "tsk_UIDPredecessors", "MKey": "tsk_MilestoneKey",
    "ACin": "tsk_ActCostInput", "Budget": "tsk_Budget", "PV": "tsk_PV", "EV": "tsk_EV", "ACt": "tsk_AC",
    "PlanPct": "tsk_PlannedPct", "LBud": "tsk_LeafBudget", "LPV": "tsk_LeafPV", "LEV": "tsk_LeafEV",
    "LAC": "tsk_LeafAC", "Phase": "tsk_Phase", "PhaseKey": "tsk_PhaseKey", "MsKey": "tsk_MilestoneOrder", "LateKey": "tsk_LateKey",
    "StartKey": "tsk_StartSoonKey", "DueKey": "tsk_DueSoonKey", "CritKey": "tsk_CriticalKey",
    "OverAlloc": "tsk_OverAllocated", "SP": "tsk_StoryPoints", "SprIn": "tsk_PlannedSprint", "LWDone": "tsk_LeafWorkDone", "DoneDate": "tsk_DoneDate",
    "LPts": "tsk_LeafPoints", "CurKey": "tsk_CurrentSprintKey",
}


def R(key):
    """This row's value of a column: [@Column]."""
    return T.row(N[key])


def C(key):
    """The whole column: Tasks[Column]."""
    return T.col(N[key])


def H(key):
    """The column's header cell: an anchor for ROW."""
    return T.head(N[key])


def other_row(key, pos):
    """This column's value in the task at position pos (1 = first task). OFFSET from the header: see module docstring."""
    return f"OFFSET({H(key)},{pos},0)"


def block(key):
    """A summary row's descendants (the rows directly below it) in one column. INDEX:INDEX is not volatile."""
    return f"INDEX({C(key)},{R('Idx')}+1):INDEX({C(key)},{R('Idx')}+{R('Desc')})"


NO_LIMIT = 2958465          # 31-Dec-9999: "no successor constrains this"


def below(key):
    """The rows from the one below this down to the last successor, in one column (successors sit below their
    predecessors). OFFSET from the header so it stays correct when rows are inserted or deleted; stopping at
    _LastSuccessorRow keeps each look-up short."""
    return f"OFFSET({H(key)},{R('Idx')}+1,0,MAX(1,{R('LastSucc')}-{R('Idx')}),1)"


SLACK_OFFSET = 100000       # link slack is stored as SLACK_OFFSET - slack (always positive)


def successor_max(key, pred_key):
    """Largest key among the rows below that name this task in pred_key; 0 when no row does.
    Keys are stored so that the largest key is the earliest date / smallest slack, and every real key is > 0:
    one MAXIFS then answers "is there a successor, and what is its tightest limit" (no COUNTIFS needed)."""
    return f'IF({R("LastSucc")}<={R("Idx")},0,_xlfn.MAXIFS({below(key)},{below(pred_key)},{R("UID")}))'


def wdiff(a, b):
    """Signed working-day difference b - a."""
    return f"IF({b}>={a},NETWORKDAYS.INTL({a},{b},{WD})-1,1-NETWORKDAYS.INTL({b},{a},{WD}))"


def pred_pos(p):
    """Position of the predecessor in the table (1 = first task; 0 = no predecessor, -1 = UID not found).
    Looked up once per link, so the columns that need the predecessor don't each repeat a MATCH."""
    return f'IF(N({R(p)})=0,0,IFERROR(MATCH({R(p)},{C("UID")},0),-1))'


def link_problem(p, pos):
    """Why a predecessor link can't be used ("" = fine). Reads no dates, so it can never be circular."""
    pred, at = R(p), R(pos)
    return (f'IF(OR(N({pred})=0,{R("Name")}=""),"",IF({R("IsSum")},"ignored: summary tasks roll up",'
            f'IF({pred}={R("UID")},"links to itself",IF({at}<0,"UID "&{pred}&" not found",'
            f'IF(INDEX({C("Name")},{at})="","UID "&{pred}&" is an empty row",'
            f'IF({at}>{R("Idx")},"UID "&{pred}&" is below this task: move it above",'
            f'IF(INDEX({C("IsSum")},{at}),"UID "&{pred}&" is a summary: link to one of its tasks","")))))))')


def driver(p, t, lag, problem, start_key, finish_key):
    """Earliest start imposed by one predecessor link (FS/SS/FF/SF + lag); 0 if there is no usable link."""
    pos = R("R1" if p == "P1" else "R2")
    pred_finish = other_row(finish_key, pos)
    pred_start = other_row(start_key, pos)
    lag, dur = f"N({R(lag)})", R("Dur")
    back = f"-MAX(0,N({dur})-1)"
    return (f'IF(OR({pos}<=0,{R(problem)}<>"",{R("IsSum")},{R("Name")}=""),0,'
            f'IFERROR(CHOOSE(MATCH(IF({R(t)}="","FS",{R(t)}),{{"FS","SS","FF","SF"}},0),'
            f'WORKDAY.INTL({pred_finish},IF(N({dur})=0,0,1)+{lag},{WD}),'
            f'WORKDAY.INTL({pred_start},{lag},{WD}),'
            f'WORKDAY.INTL(WORKDAY.INTL({pred_finish},{lag},{WD}),{back},{WD}),'
            f'WORKDAY.INTL(WORKDAY.INTL({pred_start},{lag},{WD}),{back},{WD})),0))')


def pred_text(p, t, lag, problem, by_id):
    """One link in MS Project notation: 12, 6SS+3d, 17FF+1d."""
    typ, lg = f'IF({R(t)}="","FS",{R(t)})', f"N({R(lag)})"
    ref = f"INDEX({C('ID')},{R('R1' if p == 'P1' else 'R2')})" if by_id else R(p)
    return (f'IF(OR(N({R(p)})=0,{R(problem)}<>""),"",{ref}&IF(AND({typ}="FS",{lg}=0),"",'
            f'{typ}&IF({lg}=0,"",IF({lg}>0,"+","")&{lg}&"d")))')


def formulas():
    nm, s, leaf = R("Name"), R("IsSum"), f'{R("IsLeaf")}=1'
    blank = f'{nm}=""'
    manual = f'IF(N({R("Cons")})>0,{R("Cons")},ProjectStart)'
    f = {}
    # outline
    f["ID"] = f'IF({blank},"",COUNTA({T.upto(N["Name"])})-1)'
    f["Lvl"] = f'IF({blank},"",MAX(1,N({R("Level")})))'
    f["Idx"] = f"ROW()-ROW({H('ID')})"
    f["Parent"] = (f'IF(OR({blank},{R("Lvl")}<=1),0,IFERROR(LOOKUP(2,1/({T.upto(N["Lvl"])}={R("Lvl")}-1),'
                   f'ROW({T.upto(N["Lvl"])})),0))')
    f["WBS"] = (f'IF({blank},"",IF({R("Lvl")}=1,COUNTIF({T.upto(N["Lvl"])},1)&"",IF({R("Parent")}=0,"?",'
                f'{other_row("WBS", R("Parent") + "-ROW(" + H("WBS") + ")")}&"."&'
                f'COUNTIF(INDEX({C("Lvl")},{R("Parent")}-ROW({H("Lvl")})+1):{R("Lvl")},{R("Lvl")}))))')
    # a row is a summary when the next row is one level deeper; only summaries need the descendant count
    f["NextLvl"] = f'IFERROR(N(INDEX({C("Lvl")},{R("Idx")}+1)),0)'
    f["IsSum"] = f'AND({nm}<>"",{R("NextLvl")}>{R("Lvl")})'
    f["Desc"] = f'IF({s},COUNTIF({C("WBS")},{R("WBS")}&".*"),0)'
    f["IsLeaf"] = f'IF(AND({nm}<>"",NOT({s})),1,0)'
    # links
    f["R1"] = pred_pos("P1")
    f["R2"] = pred_pos("P2")
    f["M1"] = link_problem("P1", "R1")
    f["M2"] = link_problem("P2", "R2")
    f["LinkChk"] = (f'IF({blank},"",TRIM(IF({R("M1")}<>"","Pred 1: "&{R("M1")}&". ","")&'
                    f'IF({R("M2")}<>"","Pred 2: "&{R("M2")}&".","")))')
    f["D1P"] = driver("P1", "T1", "L1", "M1", "LCS", "LCF")
    f["D2P"] = driver("P2", "T2", "L2", "M2", "LCS", "LCF")
    f["D1F"] = driver("P1", "T1", "L1", "M1", "LFS", "LFF")
    f["D2F"] = driver("P2", "T2", "L2", "M2", "LFS", "LFF")
    # detail-task dates (summaries leave these blank)
    f["LCS"] = (f'IF(NOT({leaf}),"",IF({R("Mode")}="Manual",{manual},'
                f'WORKDAY.INTL(MAX(ProjectStart,N({R("Cons")}),{R("D1P")},{R("D2P")})-1,1,{WD})))')
    f["LCF"] = f'IF(NOT({leaf}),"",IF(N({R("Dur")})<=0,{R("LCS")},WORKDAY.INTL({R("LCS")},{R("Dur")}-1,{WD})))'
    f["LFS"] = (f'IF(NOT({leaf}),"",IF(N({R("AS")})>0,{R("AS")},IF({R("Mode")}="Manual",{manual},'
                f'WORKDAY.INTL(MAX(ProjectStart,N({R("Cons")}),{R("D1F")},{R("D2F")},'
                f'IF(RescheduleAfterStatus="Yes",StatusDate,0))-1,1,{WD}))))')
    natural = f'WORKDAY.INTL({R("LFS")},{R("Dur")}-1,{WD})'
    remaining = f'WORKDAY.INTL(StatusWorkday,MAX(0,ROUNDUP({R("Dur")}*(1-MIN(1,N({R("Pct")}))),0)-1),{WD})'
    f["LFF"] = (f'IF(NOT({leaf}),"",IF(N({R("AF")})>0,{R("AF")},IF(N({R("Dur")})<=0,{R("LFS")},'
                f'IF(AND(N({R("AS")})>0,RescheduleAfterStatus="Yes"),MAX({natural},{remaining}),{natural}))))')
    f["LProg"] = f'IF(NOT({leaf}),0,IF(N({R("AF")})>0,1,MIN(1,N({R("Pct")}))))'
    f["LWork"] = (f'IF(OR(NOT({leaf}),{R("Res")}=""),0,'
                  f'N({R("Dur")})*HoursPerDay*IF({R("Units")}="",DefaultUnits,{R("Units")}))')
    f["LCost"] = f'IF({R("LWork")}=0,0,{R("LWork")}*IFERROR(INDEX(res_Rate,MATCH({R("Res")},res_Name,0)),0))'
    # critical path: backward pass. Each successor row works out, per link, the latest its predecessor may finish
    # (FS / FF directly; SS / SF via the predecessor's duration, an input) and stores it as a key; the predecessor
    # then takes one MAXIFS over the rows below. A task's late finish is the earliest of those limits, its deadline
    # and the project's forecast finish.
    for k in ("1", "2"):
        p, t, lag, m, d = f"P{k}", f"T{k}", f"L{k}", f"M{k}", f"D{k}F"
        valid = f'AND({leaf},{R("R" + k)}>0,{R(m)}="")'
        typ, lg = f'IF({R(t)}="","FS",{R(t)})', f"N({R(lag)})"
        pred_dur = f"N(INDEX({C('Dur')},{R('R' + k)}))"
        start_limit = f'WORKDAY.INTL(IF({typ}="SS",{R("LLS")},{R("LLF")}),-{lg},{WD})'
        finish_limit = (f'IF({typ}="FS",WORKDAY.INTL({R("LLS")},-(IF(N({R("Dur")})=0,0,1)+{lg}),{WD}),'
                        f'IF({typ}="FF",WORKDAY.INTL({R("LLF")},-{lg},{WD}),'
                        f'IF({pred_dur}<=0,{start_limit},WORKDAY.INTL({start_limit},{pred_dur}-1,{WD}))))')
        f[f"F{k}"] = f'IF({valid},{NO_LIMIT}-{finish_limit},"")'
        f[f"K{k}"] = f'IF({valid},{SLACK_OFFSET}-{wdiff(R(d), R("LFS"))},"")'
    # where the last successor sits: reads only input columns, so it isn't recalculated after every edit
    f["LastSucc"] = (f'IF(NOT({leaf}),0,MAX(_xlfn.MAXIFS({C("Idx")},{C("P1")},{R("UID")}),'
                     f'_xlfn.MAXIFS({C("Idx")},{C("P2")},{R("UID")})))')
    succ = f'MAX({successor_max("F1", "P1")},{successor_max("F2", "P2")})'
    f["SuccLF"] = f'IF(NOT({leaf}),"",{NO_LIMIT}-{succ})'
    f["LLF"] = (f'IF(NOT({leaf}),"",MIN(ProjectFinish,IF(N({R("DL")})>0,{R("DL")},{NO_LIMIT}),{R("SuccLF")}))')
    f["LLS"] = f'IF(NOT({leaf}),"",IF(N({R("Dur")})<=0,{R("LLF")},WORKDAY.INTL({R("LLF")},-({R("Dur")}-1),{WD})))'
    f["LTS"] = f'IF(NOT({leaf}),"",{wdiff(R("LFF"), R("LLF"))})'
    f["LCrit"] = f'AND({leaf},{R("LProg")}<1,N({R("LTS")})<=CriticalSlackDays)'
    slack = f'MAX({successor_max("K1", "P1")},{successor_max("K2", "P2")})'
    f["FSl"] = f'IF(NOT({leaf}),"",IF({slack}=0,{R("LTS")},MIN({R("LTS")},{SLACK_OFFSET}-{slack})))'
    f["LS"] = f'IF({blank},"",IF({s},MIN({block("LLS")}),{R("LLS")}))'
    f["LF"] = f'IF({blank},"",IF({s},MAX({block("LLF")}),{R("LLF")}))'
    f["TS"] = f'IF({blank},"",IF({s},MIN({block("LTS")}),{R("LTS")}))'
    f["Crit"] = (f'IF({blank},"",IF(IF({s},COUNTIF({block("LCrit")},TRUE())>0,{R("LCrit")}),"Yes",""))')
    # every row: detail value, or the roll-up of the detail tasks below a summary
    rollup = lambda fn, leaf_key: f'IF({blank},"",IF({s},{fn}({block(leaf_key)}),{R(leaf_key)}))'
    f["CPS"] = rollup("MIN", "LCS")
    f["CPF"] = rollup("MAX", "LCF")
    f["FS"] = rollup("MIN", "LFS")
    f["FF"] = rollup("MAX", "LFF")
    f["Work"] = rollup("SUM", "LWork")
    f["Cost"] = rollup("SUM", "LCost")
    f["Prog"] = (f'IF({blank},"",IF({s},IFERROR(SUM({block("WtDone")})/SUM({block("WtDur")}),0),{R("LProg")}))')
    f["PS"] = f'IF({blank},"",IF(N({R("BS")})>0,{R("BS")},{R("CPS")}))'
    f["PF"] = f'IF({blank},"",IF(N({R("BF")})>0,{R("BF")},{R("CPF")}))'
    f["SVar"] = f'IF({blank},"",{wdiff(R("PS"), R("FS"))})'
    f["Var"] = f'IF({blank},"",{wdiff(R("PF"), R("FF"))})'
    f["DLChk"] = (f'IF(OR({blank},N({R("DL")})=0),"",IF({R("FF")}>{R("DL")},'
                  f'"Late by "&({wdiff(R("DL"), R("FF"))})&"d","OK"))')
    f["Status"] = (f'IF({blank},"",IF({R("Prog")}>=1,"Complete",IF(OR(N({R("AS")})>0,{R("Prog")}>0),'
                   f'IF({R("Var")}>0,"Behind","In Progress"),IF({R("PS")}<StatusDate,"Late Start","Not Started"))))')
    f["Sprint"] = (f'IF(OR({blank},{s}),"",IF({R("SprIn")}<>"",{R("SprIn")},'
                   f'IF(OR({R("FS")}<Sprint1Start,{R("FS")}>SprintLastEnd),"Backlog",'
                   f'INDEX(spr_Name,MATCH({R("FS")},spr_Start,1)))))')
    f["T1x"] = pred_text("P1", "T1", "L1", "M1", True)
    f["T2x"] = pred_text("P2", "T2", "L2", "M2", True)
    f["U1x"] = pred_text("P1", "T1", "L1", "M1", False)
    f["U2x"] = pred_text("P2", "T2", "L2", "M2", False)
    join = lambda a, b: f'IF({blank},"",{R(a)}&IF(AND({R(a)}<>"",{R(b)}<>""),",","")&{R(b)})'
    f["PredTxt"] = join("T1x", "T2x")
    f["UPredTxt"] = join("U1x", "U2x")
    f["SrcS"] = f'IF({blank},"",{R("FS")})'
    f["SrcF"] = f'IF({blank},"",{R("FF")})'
    f["SrcD"] = f'IF({blank},"",IF({s},NETWORKDAYS.INTL({R("FS")},{R("FF")},{WD}),N({R("Dur")})))'
    f["SrcW"] = f'IF({blank},"",{R("Work")})'
    f["SrcC"] = f'IF({blank},"",{R("Cost")})'
    variance = lambda cur, base: f'IF(OR({blank},{R(base)}=""),"",{R(cur)}-N({R(base)}))'
    f["DVar"] = variance("SrcD", "BD")
    f["WVar"] = variance("Work", "BW")
    f["CVar"] = variance("Cost", "BC")
    # time-phased work, in working days counted from the project start (used by Sprints with SUMIFS):
    # work up to working day D = SUMIFS(work, finish < D) + D x SUMIFS(work/day, running on D) - SUMIFS(work before start)
    has_work = f'AND({leaf},{R("LWork")}>0)'
    f["SWD"] = f'IF({has_work},NETWORKDAYS.INTL(ProjectStart,{R("LFS")},{WD}),0)'
    f["FWD"] = f'IF({has_work},MAX({R("SWD")},NETWORKDAYS.INTL(ProjectStart,{R("LFF")},{WD})),0)'
    f["WPD"] = f'IF({has_work},{R("LWork")}/({R("FWD")}-{R("SWD")}+1),0)'
    f["WBefore"] = f'{R("WPD")}*({R("SWD")}-1)'
    # progress roll-up and the milestone list
    f["WtDur"] = f'IF({leaf},MAX(N({R("Dur")}),0.1),0)'
    f["WtDone"] = f'{R("WtDur")}*{R("LProg")}'
    f["MKey"] = f'IF(AND({leaf},N({R("Dur")})=0,{R("LProg")}<1),{R("LFF")}+{R("Idx")}/100000,"")'
    # earned value at the status date. Budget = Baseline Cost once a baseline is set, otherwise Cost. Planned % =
    # share of the planned working days before the status date (the status date itself is not yet worked).
    # Actual cost = the Actual Cost entered, or Cost x Progress when it is blank (as MS Project calculates it).
    ps, pf = R("PS"), R("PF")
    f["PlanPct"] = (f'IF(NOT({leaf}),0,IF({pf}<StatusDate,1,IF({ps}>=StatusDate,0,'
                    f'IFERROR(MIN(1,NETWORKDAYS.INTL({ps},StatusDate-1,{WD})/NETWORKDAYS.INTL({ps},{pf},{WD})),0))))')
    f["LBud"] = f'IF(NOT({leaf}),0,IF({R("BC")}="",{R("LCost")},N({R("BC")})))'
    f["LPV"] = f'{R("LBud")}*{R("PlanPct")}'
    f["LEV"] = f'{R("LBud")}*{R("LProg")}'
    f["LAC"] = f'IF(NOT({leaf}),0,IF({R("ACin")}="",{R("LCost")}*{R("LProg")},N({R("ACin")})))'
    f["Budget"] = rollup("SUM", "LBud")
    f["PV"] = rollup("SUM", "LPV")
    f["EV"] = rollup("SUM", "LEV")
    f["ACt"] = rollup("SUM", "LAC")
    # report lists (Schedule Report): each key is unique (the row position breaks ties) and blank when the task
    # doesn't belong on the list, so the report finds the n-th task with SMALL / LARGE and MATCH
    tie = f'{R("Idx")}/100000'
    f["Phase"] = f'IF({blank},"",LEFT({R("WBS")},FIND(".",{R("WBS")}&".")-1))'
    f["PhaseKey"] = f'IF(AND({nm}<>"",{R("Lvl")}=1),{R("Idx")},"")'
    f["MsKey"] = f'IF(AND({leaf},N({R("Dur")})=0),{pf}+{tie},"")'
    f["LateKey"] = (f'IF(AND({leaf},{R("LProg")}<1,OR(N({R("Var")})>0,{R("Status")}="Late Start")),'
                    f'MAX(0,{R("Var")})+1-{R("Idx")}/1000000,"")')
    f["StartKey"] = (f'IF(AND({leaf},N({R("AS")})=0,{R("LFS")}>=StatusDate,{R("LFS")}<=LookaheadEnd),'
                     f'{R("LFS")}+{tie},"")')
    f["DueKey"] = (f'IF(AND({leaf},{R("LProg")}<1,{R("LFF")}>=StatusDate,{R("LFF")}<=LookaheadEnd),'
                   f'{R("LFF")}+{tie},"")')
    f["CritKey"] = f'IF({R("LCrit")},{R("LFS")}+{tie},"")'
    # resource and agile reports. A task is done on its Actual Finish (or its forecast finish when it is marked
    # 100% without one); story points count on detail tasks only.
    f["LWDone"] = f'{R("LWork")}*{R("LProg")}'
    f["DoneDate"] = f'IF(AND({leaf},{R("LProg")}>=1),IF(N({R("AF")})>0,{R("AF")},{R("LFF")}),"")'
    f["LPts"] = f'IF({leaf},N({R("SP")}),0)'
    # over-allocation (Resource Report weekly utilisation, after leave): Yes when the task's resource is over 100% in
    # any week the task runs. Only weeks in the Resource Report window and its first resources are known.
    week_of = lambda d: f"(INT(({d}-ResourceWeekStart)/7)+1)"
    first, last = f"MAX(1,{week_of(R('LFS'))})", f"MIN(COLUMNS(rw_Util),{week_of(R('LFF'))})"
    util_row = f"MATCH({R('Res')},res_Name,0)"
    weeks = f"INDEX(rw_Util,{util_row},{first}):INDEX(rw_Util,{util_row},{last})"
    f["OverAlloc"] = (f'IF(OR(NOT({leaf}),{R("Res")}="",{R("LProg")}>=1),"",IFERROR(IF({first}>{last},"",'
                      f'IF(COUNTIF({weeks},">1")>0,"Yes","")),""))')
    f["CurKey"] = f'IF(AND({leaf},CurrentSprint<>"",{R("Sprint")}=CurrentSprint),{R("Idx")},"")'
    return f


# key, kind, width, number format, centre, header note
LAYOUT = [
    ("ID", "calc", 5, None, True, "Position in the list; renumbers when rows move (like MS Project)."),
    ("UID", "in", 6, None, True, "Permanent number that predecessors refer to. Each task needs a different one."),
    ("WBS", "calc", 8, None, True, "Outline number, generated from Level."),
    ("Level", "in", 6, None, True, "Outline level: 1 = phase, 2 = task, 3 = sub-task ... (at most 1 more than the row above)."),
    ("Name", "in", 36, None, False, None),
    ("Mode", "in", 8, None, True, "Auto = scheduled from predecessors. Manual = pinned to Start No Earlier Than."),
    ("Dur", "in", 9, None, True, "Working days. 0 = milestone."),
    ("P1", "in", 7, None, True, "UID of the predecessor."),
    ("T1", "in", 6, None, True, "FS, SS, FF or SF (blank = FS)."),
    ("L1", "in", 6, None, True, "Lag in working days (negative = lead)."),
    ("P2", "in", 7, None, True, "UID of a second predecessor."),
    ("T2", "in", 6, None, True, "FS, SS, FF or SF (blank = FS)."),
    ("L2", "in", 6, None, True, "Lag in working days (negative = lead)."),
    ("Cons", "in", 15, DATE, False, "Auto tasks start no earlier than this. Manual tasks start exactly on it."),
    ("DL", "in", 15, DATE, False, "Optional deadline; Deadline Check flags a forecast finish after it."),
    ("PS", "calc", 15, DATE, False, "Baseline Start if a baseline is set, otherwise calculated from durations and links."),
    ("PF", "calc", 15, DATE, False, "Baseline Finish if a baseline is set, otherwise calculated."),
    ("AS", "in", 15, DATE, False, None),
    ("AF", "in", 15, DATE, False, None),
    ("FS", "calc", 15, DATE, False, "Current schedule: actuals where entered, otherwise recalculated from predecessors."),
    ("FF", "calc", 15, DATE, False, "Current schedule finish."),
    ("SVar", "calc", 8, None, True, "Forecast Start minus Planned Start, in working days."),
    ("Var", "calc", 8, None, True, "Forecast Finish minus Planned Finish, in working days."),
    ("LS", "calc", 15, DATE, False, "Latest start that doesn't delay the project finish or a deadline (critical path backward pass)."),
    ("LF", "calc", 15, DATE, False, "Latest finish that doesn't delay the project finish or a deadline."),
    ("TS", "calc", 8, None, True, "Total slack: working days this task can slip before it delays the project or a deadline."),
    ("FSl", "calc", 8, None, True, "Free slack: working days it can slip before it delays any successor."),
    ("Crit", "calc", 8, None, True, "Yes when Total Slack <= CriticalSlackDays (Settings) and the task isn't complete."),
    ("DLChk", "calc", 12, None, True, None),
    ("LinkChk", "calc", 24, None, False, "Explains predecessor links that are ignored."),
    ("Pct", "in", 9, "0%", False, "Enter 0-100%. An Actual Finish counts as 100%."),
    ("Prog", "calc", 9, "0%", False, "% Complete, rolled up to summaries (duration-weighted)."),
    ("Status", "calc", 12, None, True, None),
    ("Res", "in", 16, None, False, None),
    ("Units", "in", 7, "0%", False, "Share of the resource's time (blank = DefaultUnits)."),
    ("OverAlloc", "calc", 9, None, True, "Yes when this task's resource is over 100% of capacity (after leave) in a "
                                         "week the task runs: see Resource Report. Only weeks shown there count."),
    ("Work", "calc", 9, "#,##0.0", False, "Hours = Duration x HoursPerDay x Units."),
    ("Cost", "calc", 11, "#,##0.00", False, None),
    ("ACin", "in", 11, "#,##0.00", False, "Cost to date of a detail task. Blank = Cost x Progress (as MS Project "
                                          "calculates it). Summaries roll up in the AC column."),
    ("SprIn", "in", 11, None, True, "Optional: the sprint this task is planned in (dropdown). Blank = the sprint "
                                      "containing its forecast start."),
    ("Sprint", "calc", 10, None, True, "Planned Sprint if entered, otherwise the sprint containing the forecast start."),
    ("SP", "in", 8, "0.#", True, "Story points (optional estimate for agile reporting). Detail tasks only."),
    ("PredTxt", "calc", 14, None, False, "Predecessors in MS Project notation, by ID."),
    ("Notes", "in", 40, None, False, None),
    ("SrcS", "calc", 15, DATE, False, "Copy the 5 BL Source columns, then Paste Special > Values onto the 5 Baseline columns."),
    ("SrcF", "calc", 15, DATE, False, None),
    ("SrcD", "calc", 9, None, True, None),
    ("SrcW", "calc", 9, "#,##0.0", False, None),
    ("SrcC", "calc", 11, "#,##0.00", False, None),
    ("BS", "in", 15, DATE, False, "Frozen plan. Filled by pasting values from the BL Source columns."),
    ("BF", "in", 15, DATE, False, None),
    ("BD", "in", 9, None, True, None),
    ("BW", "in", 9, "#,##0.0", False, None),
    ("BC", "in", 11, "#,##0.00", False, None),
    ("DVar", "calc", 9, None, True, "Duration minus Baseline Duration (working days)."),
    ("WVar", "calc", 9, "#,##0.0", False, "Work minus Baseline Work (hours)."),
    ("CVar", "calc", 11, "#,##0.00", False, "Cost minus Baseline Cost."),
    ("Budget", "calc", 11, "#,##0.00", False, "Budget at completion: Baseline Cost once a baseline is set, otherwise Cost."),
    ("PV", "calc", 11, "#,##0.00", False, "Planned value: the budget of the work planned before the status date."),
    ("EV", "calc", 11, "#,##0.00", False, "Earned value: Budget x Progress."),
    ("ACt", "calc", 11, "#,##0.00", False, "Actual cost to date: Actual Cost, or Cost x Progress when it is blank."),
] + [(k, "help", 11, DATE if k in ("CPS", "CPF", "LCS", "LCF", "LFS", "LFF", "D1P", "D2P", "D1F", "D2F", "SuccLF",
                                       "LLF", "LLS", "DoneDate")
                          else None, False, None)
     for k in ("Lvl", "Idx", "Parent", "NextLvl", "IsSum", "Desc", "IsLeaf", "R1", "R2", "LCS", "LCF", "LFS", "LFF",
               "LProg", "LWork", "LCost", "F1", "F2", "K1", "K2", "LastSucc", "SuccLF", "LLF", "LLS",
               "LTS", "LCrit", "CPS", "CPF", "M1", "M2", "D1P", "D2P", "D1F",
               "D2F", "SWD", "FWD", "WPD", "WBefore", "WtDur", "WtDone", "T1x", "T2x", "U1x", "U2x", "UPredTxt", "MKey",
               "PlanPct", "LBud", "LPV", "LEV", "LAC", "Phase", "PhaseKey", "MsKey", "LateKey", "StartKey", "DueKey", "CritKey",
               "LWDone", "DoneDate", "LPts", "CurKey")]

GROUPS = [("ID", "IDENTITY"), ("Mode", "SCHEDULING INPUTS"), ("PS", "PLAN / ACTUAL / FORECAST"), ("SVar", "VARIANCE"),
          ("LS", "CRITICAL PATH"), ("DLChk", "CHECKS"),
          ("Pct", "PROGRESS"), ("Res", "RESOURCES & SPRINT"),
          ("SrcS", "SET BASELINE: copy these 5 columns, then Paste Special > Values onto the 5 Baseline columns"),
          ("DVar", "VARIANCE vs BASELINE"), ("Budget", "EARNED VALUE (at status date)"), ("Lvl", "HELPER COLUMNS (hidden - click + above to show)")]
STATUS_COLOURS = {"Complete": "548235", "In Progress": "2F75B5", "Behind": "C00000", "Late Start": "C55A11",
                  "Not Started": "7F7F7F"}


def build_tasks(wb, rows, n_rows):
    ws = wb.create_sheet("Tasks")
    title(ws, None)
    ws["A1"] = '=ProjectName&"  -  Task Entry"'
    ws["A2"] = ("Type in BLUE columns only. Level = outline level (1 = phase, 2 = task, 3 = sub-task). Predecessors refer "
                "to the UID column. Planned = Baseline once set, otherwise calculated from Duration + Predecessors. "
                "Forecast also uses Actuals, so slips ripple to every successor. Insert rows inside the table: formulas "
                "fill in automatically.")
    ws["A2"].font = NOTE_FONT
    ws["A3"] = '="Next free UID for a new row: "&NextFreeUID'
    ws["A3"].font = LABEL_FONT
    f = formulas()
    cols = []
    for key, kind, width, fmt, center, note in LAYOUT:
        values = []
        if key == "UID":
            values = list(range(1, n_rows + 1))           # stable IDs: links keep working when rows move
        elif key in SAMPLE_COLUMNS:
            values = [r[key] for r in rows]
        cols.append(Col(N[key], kind, width, fmt, None if kind == "in" else f[key], center, note, key, values))
    L = add_table(ws, TABLE, HEADER_ROW, 1, cols, n_rows)
    ws.row_dimensions[HEADER_ROW].height = 42
    for key, label in GROUPS:
        ws[f"{L[key]}4"] = label
        ws[f"{L[key]}4"].font = LABEL_FONT
    for key, defined in NAMES.items():
        name(wb, defined, C(key))
    ws.freeze_panes = f"{L['Mode']}{HEADER_ROW + 1}"
    ws.column_dimensions.group(L["Lvl"], L["CurKey"], hidden=True, outline_level=1)
    _validation_and_formatting(ws, L, n_rows)
    ws.sheet_view.zoomScale = 90
    return ws


def _validation_and_formatting(ws, L, n_rows):
    """Data validation and conditional formatting: Excel only accepts A1 ranges here (it shifts them on insert)."""
    first, last = HEADER_ROW + 1, HEADER_ROW + n_rows
    rng = lambda k: col_range(L[k], first, last)
    at = lambda k: f"${L[k]}{first}"           # this row's cell, relative to the top-left of the formatted range

    def dv(rule, *keys):
        ws.add_data_validation(rule)
        for k in keys:
            rule.add(rng(k))

    dv(DataValidation(type="whole", operator="greaterThanOrEqual", formula1="1", showErrorMessage=True, errorTitle="UID",
                      error="UID must be a whole number; give each task a different one (see Settings > NextFreeUID)."), "UID")
    dv(DataValidation(type="whole", operator="between", formula1="1", formula2="9", showErrorMessage=True,
                      errorTitle="Level", error="Outline level must be 1-9 and at most 1 more than the row above."), "Level")
    dv(DataValidation(type="list", formula1="=L_Mode"), "Mode")
    dv(DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", showErrorMessage=True,
                      error="Duration is in working days (0 = milestone)."), "Dur")
    dv(DataValidation(type="whole", operator="greaterThanOrEqual", formula1="1", showErrorMessage=True,
                      error="Enter the UID of another task. Problems are explained in the Link Check column."), "P1", "P2")
    dv(DataValidation(type="list", formula1="=L_DepType"), "T1", "T2")
    dv(DataValidation(type="whole", operator="between", formula1="-365", formula2="365"), "L1", "L2")
    dv(DataValidation(type="date", operator="greaterThan", formula1="1"), "Cons", "DL", "AS", "AF")
    dv(DataValidation(type="decimal", operator="between", formula1="0", formula2="1", showErrorMessage=True,
                      error="Enter 0% - 100%."), "Pct")
    dv(DataValidation(type="list", formula1="=res_Name"), "Res")
    dv(DataValidation(type="decimal", operator="between", formula1="0", formula2="10"), "Units")
    dv(DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0"), "ACin", "SP")
    dv(DataValidation(type="list", formula1="=spr_Name"), "SprIn")

    cf = ws.conditional_formatting
    cf.add(f"A{first}:{L['ACt']}{last}", FormulaRule(formula=[f"{at('IsSum')}=TRUE"], font=Font(bold=True), fill=fill("DDEBF7")))
    for k in ("Dur", "P1", "T1", "L1", "P2", "T2", "L2", "Cons", "AS", "AF", "Pct", "Res", "Units", "ACin", "SprIn", "SP"):
        cf.add(rng(k), FormulaRule(formula=[f"{at('IsSum')}=TRUE"], font=Font(color="A6A6A6", italic=True, bold=True)))
    for status, colour in STATUS_COLOURS.items():
        cf.add(rng("Status"), FormulaRule(formula=[f'{at("Status")}="{status}"'], font=Font(color=colour, bold=True)))
    for k in ("SVar", "Var", "DVar", "WVar", "CVar"):
        cf.add(rng(k), FormulaRule(formula=[f"AND(ISNUMBER({at(k)}),{at(k)}>0)"], font=Font(color="C00000", bold=True)))
        cf.add(rng(k), FormulaRule(formula=[f"AND(ISNUMBER({at(k)}),{at(k)}<0)"], font=Font(color="548235", bold=True)))
    cf.add(rng("Name"), FormulaRule(formula=[f"AND({at('IsLeaf')}=1,N({at('Dur')})=0)"], font=Font(italic=True, color="7030A0")))
    cf.add(rng("WBS"), FormulaRule(formula=[f'{at("WBS")}="?"'], fill=fill("FFC7CE")))
    cf.add(rng("UID"), FormulaRule(
        formula=[f'OR(AND({at("Name")}<>"",{at("UID")}=""),AND({at("UID")}<>"",COUNTIF(tsk_UID,{at("UID")})>1))'],
        fill=fill("FFC7CE"), font=Font(color="9C0006", bold=True)))
    cf.add(rng("LinkChk"), FormulaRule(formula=[f'{at("LinkChk")}<>""'], fill=fill("FFC7CE"), font=Font(color="9C0006")))
    cf.add(rng("DLChk"), FormulaRule(formula=[f'LEFT({at("DLChk")},4)="Late"'], font=Font(color="C00000", bold=True)))
    cf.add(rng("OverAlloc"), FormulaRule(formula=[f'{at("OverAlloc")}="Yes"'], font=Font(color="9C0006", bold=True),
                                         fill=fill("FFC7CE")))
    cf.add(rng("Crit"), FormulaRule(formula=[f'{at("Crit")}="Yes"'], font=Font(color="FFFFFF", bold=True), fill=fill("C00000")))
    for k in ("TS", "FSl"):
        cf.add(rng(k), FormulaRule(formula=[f"AND(ISNUMBER({at(k)}),{at(k)}<0)"], font=Font(color="C00000", bold=True)))
    cf.add(rng("DLChk"), FormulaRule(formula=[f'{at("DLChk")}="OK"'], font=Font(color="548235", bold=True)))
