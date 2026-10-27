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
import copy

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.cell_range import MultiCellRange
from openpyxl.worksheet.datavalidation import DataValidation

from .core import (DATE, LABEL_FONT, NOTE_FONT, Col, Tab, add_table, col_range, fill, name, title)

TABLE = "Tasks"
HEADER_ROW = 5
T = Tab(TABLE)
WD = "WorkWeek,Holidays"
LINKS = 4                   # predecessor links per task: Pred1 ... Pred4 (raise it here; every link column follows)
LINK_NUMBERS = range(1, LINKS + 1)
PEOPLE = 4                  # people per task in the Resources cell; the first uses the Resource / _Leaf... columns
MORE_PEOPLE = range(2, PEOPLE + 1)
# per-person values other sheets total by resource: (first person's defined name, name for person k)
PERSON_FIELDS = {"Name": ("tsk_Resource", "tsk_Res{k}"), "Work": ("tsk_LeafWork", "tsk_Res{k}Work"),
                 "Cost": ("tsk_LeafCost", "tsk_Res{k}Cost"), "AC": ("tsk_ResourceAC", "tsk_Res{k}AC"),
                 "Done": ("tsk_LeafWorkDone", "tsk_Res{k}Done"), "WPD": ("tsk_WorkPerWD", "tsk_Res{k}WPD"),
                 "WBefore": ("tsk_WorkBeforeStart", "tsk_Res{k}WBefore")}


def people(field):
    """The defined names holding one per-person value, first person first: people("Work") = [tsk_LeafWork, ...]."""
    first, more = PERSON_FIELDS[field]
    return [first] + [more.format(k=k) for k in MORE_PEOPLE]
# MS Project's constraint types, in the order of its ConstraintType codes (0 = As Soon As Possible ... 7)
CONSTRAINTS = ["As Soon As Possible", "As Late As Possible", "Must Start On", "Must Finish On", "Start No Earlier Than",
               "Start No Later Than", "Finish No Earlier Than", "Finish No Later Than"]

# key -> column name (= header text = structured-reference name)
N = {
    "ID": "ID", "UID": "UID", "WBS": "WBS", "Level": "Level", "Show": "Show", "Name": "Task Name", "Mode": "Mode", "Dur": "Duration",
    "PredIn": "Predecessors",
    **{f"{x}{k}": f"_Pred{k}{label}" for k in LINK_NUMBERS for x, label in (("I", "Text"), ("P", "UID"), ("T", "Type"),
                                                                             ("L", "Lag"))},
    "CType": "Constraint", "Cons": "Constraint Date", "DL": "Deadline",
    "PS": "Planned Start", "PF": "Planned Finish", "AS": "Actual Start", "AF": "Actual Finish",
    "FS": "Forecast Start", "FF": "Forecast Finish", "SVar": "Start Var", "Var": "Finish Var",
    "LS": "Late Start", "LF": "Late Finish", "TS": "Total Slack", "FSl": "Free Slack", "Crit": "Critical",
    "DLChk": "Deadline Check", "LinkChk": "Link Check", "Issues": "Issues", "Pct": "% Complete", "Prog": "Progress", "Status": "Status",
    "ResIn": "Resources", "Res": "_Resource1", "Units": "_Units1", "OverAlloc": "Over-allocated", "Work": "Work", "Cost": "Cost", "ACin": "Actual Cost", "SprIn": "Planned Sprint",
    "Sprint": "Sprint", "SP": "Story Points",
    "PredTxt": "Predecessors by ID", "Notes": "Notes",
    "SrcS": "BL Source Start", "SrcF": "BL Source Finish", "SrcD": "BL Source Duration", "SrcW": "BL Source Work",
    "SrcC": "BL Source Cost", "BS": "Baseline Start", "BF": "Baseline Finish", "BD": "Baseline Duration",
    "BW": "Baseline Work", "BC": "Baseline Cost", "DVar": "Duration Var", "WVar": "Work Var", "CVar": "Cost Var",
    "Budget": "Budget", "PV": "PV", "EV": "EV", "ACt": "AC",
    # helper columns (hidden)
    "Lvl": "_Level", "Idx": "_Row", "Parent": "_ParentRow", "NextLvl": "_NextLevel", "IsSum": "_IsSummary",
    "Desc": "_Descendants", "IsLeaf": "_IsLeaf",
    **{f"{x}{k}": f"_Pred{k}{label}" for k in LINK_NUMBERS for x, label in (
        ("R", "Pos"), ("M", "Problem"), ("D", "PlanDriver"), ("E", "FcstDriver"), ("F", "FinishKey"), ("K", "SlackKey"))},
    "CT": "_ConstraintCode", "CStart": "_ConstraintStart", "CLow": "_ConstraintLow", "CHigh": "_ConstraintHigh",
    "LCS": "_LeafCalcStart", "LCF": "_LeafCalcFinish", "LFS": "_LeafFcstStart", "LFF": "_LeafFcstFinish",
    "LProg": "_LeafProgress", "LWork": "_LeafWork", "LCost": "_LeafCost", "LTW": "_LeafTotalWork",
    "LTC": "_LeafTotalCost", "PAC": "_ResourceAC", "ResTxt": "Resource Names", "NPeople": "_People",
    **{f"Ri{k}": f"_Res{k}Text" for k in range(1, PEOPLE + 1)},
    **{f"{x}{k}": f"_Res{k}{label}" for k in MORE_PEOPLE for x, label in (
        ("Rs", ""), ("Ru", "Units"), ("Rw", "Work"), ("Rc", "Cost"), ("Rp", "WorkPerWD"), ("Rb", "WorkBefore"),
        ("Rd", "WorkDone"), ("Ra", "AC"))},
    "SuccLF": "_SuccessorFinishLimit", "LastSucc": "_LastSuccessorRow",
    "LLF": "_LeafLateFinish", "LLS": "_LeafLateStart", "LTS": "_LeafTotalSlack", "LCrit": "_LeafCritical",
    "CPS": "_CalcStart", "CPF": "_CalcFinish",
    "SWD": "_StartWorkday", "FWD": "_FinishWorkday", "WPD": "_WorkPerWorkday", "WBefore": "_WorkBeforeStart",
    "WtDur": "_WeightDur", "WtDone": "_WeightDone",
    "Shift": "_AlapShift", "XCS": "_ShownCalcStart", "XCF": "_ShownCalcFinish", "XFS": "_ShownFcstStart",
    "XFF": "_ShownFcstFinish", "UPredTxt": "_UIDPredecessors", "MKey": "_MilestoneKey",
    # earned value and the report lists
    "PlanPct": "_PlannedPct", "LBud": "_LeafBudget", "LPV": "_LeafPV", "LEV": "_LeafEV", "LAC": "_LeafAC",
    "Phase": "_Phase", "PhaseKey": "_PhaseKey", "MsKey": "_MilestoneOrder", "LateKey": "_LateKey", "StartKey": "_StartSoonKey",
    "DueKey": "_DueSoonKey", "CritKey": "_CriticalKey",
    # resource and agile reports
    "LWDone": "_LeafWorkDone", "DoneDate": "_DoneDate", "LPts": "_LeafPoints", "CurKey": "_CurrentSprintKey",
}

# defined names used by every other sheet (tsk_Duration = Tasks[Duration], ...)
NAMES = {
    "ID": "tsk_ID", "UID": "tsk_UID", "WBS": "tsk_WBS", "Lvl": "tsk_Level", "Show": "tsk_Show", "Name": "tsk_Name", "Mode": "tsk_Mode",
    "Dur": "tsk_Duration", "CType": "tsk_ConstraintType", "CT": "tsk_ConstraintCode", "Cons": "tsk_Constraint", "DL": "tsk_Deadline", "PS": "tsk_PlanStart",
    "PF": "tsk_PlanFinish", "AS": "tsk_ActStart", "AF": "tsk_ActFinish", "FS": "tsk_FcStart", "FF": "tsk_FcFinish",
    "SVar": "tsk_StartVariance", "Var": "tsk_Variance", "LS": "tsk_LateStart", "LF": "tsk_LateFinish",
    "TS": "tsk_TotalSlack", "FSl": "tsk_FreeSlack", "Crit": "tsk_Critical", "DLChk": "tsk_DeadlineCheck", "LinkChk": "tsk_LinkCheck", "Issues": "tsk_Issues",
    "Pct": "tsk_PctInput", "Prog": "tsk_Progress", "Status": "tsk_Status", "Res": "tsk_Resource",
    "Units": "tsk_Units", "Work": "tsk_Work", "Cost": "tsk_Cost", "Sprint": "tsk_Sprint",
    "PredTxt": "tsk_Predecessors", "PredIn": "tsk_PredecessorsInput", "Notes": "tsk_Notes", "BS": "tsk_BaseStart", "BF": "tsk_BaseFinish",
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
    "LPts": "tsk_LeafPoints", "CurKey": "tsk_CurrentSprintKey", "LFF": "tsk_LeafFcstFinish",
    "LProg": "tsk_LeafProgress", "LCost": "tsk_LeafCost", "LTC": "tsk_LeafTotalCost", "PAC": "tsk_ResourceAC",
    "ResTxt": "tsk_ResourceNames", "ResIn": "tsk_ResourcesInput", "NPeople": "tsk_People",
    **{f"{x}{k}": f"tsk_Res{k}{label}" for k in MORE_PEOPLE for x, label in (
        ("Rs", ""), ("Rw", "Work"), ("Rc", "Cost"), ("Rp", "WPD"), ("Rb", "WBefore"), ("Rd", "Done"), ("Ra", "AC"))},
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


def driver(k, start_key, finish_key):
    """Earliest start imposed by predecessor link k (FS/SS/FF/SF + lag); 0 if there is no usable link."""
    pos = R(f"R{k}")
    pred_finish = other_row(finish_key, pos)
    pred_start = other_row(start_key, pos)
    t, lag, dur = R(f"T{k}"), f"N({R(f'L{k}')})", R("Dur")
    back = f"-MAX(0,N({dur})-1)"
    return (f'IF(OR({pos}<=0,{R(f"M{k}")}<>"",{R("IsSum")},{R("Name")}=""),0,'
            f'IFERROR(CHOOSE(MATCH(IF({t}="","FS",{t}),{{"FS","SS","FF","SF"}},0),'
            f'WORKDAY.INTL({pred_finish},IF(N({dur})=0,0,1)+{lag},{WD}),'
            f'WORKDAY.INTL({pred_start},{lag},{WD}),'
            f'WORKDAY.INTL(WORKDAY.INTL({pred_finish},{lag},{WD}),{back},{WD}),'
            f'WORKDAY.INTL(WORKDAY.INTL({pred_start},{lag},{WD}),{back},{WD})),0))')


def pred_text(k, by_id):
    """Link k in MS Project notation with a leading comma (",12", ",6SS+3d"); "" when there is no usable link."""
    typ, lg = f'IF({R(f"T{k}")}="","FS",{R(f"T{k}")})', f"N({R(f'L{k}')})"
    ref = f"INDEX({C('ID')},{R(f'R{k}')})" if by_id else R(f"P{k}")
    return (f'IF(OR(N({R(f"P{k}")})=0,{R(f"M{k}")}<>""),"",","&{ref}&IF(AND({typ}="FS",{lg}=0),"",'
            f'{typ}&IF({lg}=0,"",IF({lg}>0,"+","")&{lg}&"d")))')


def links_text(links):
    """The Predecessors cell for (UID, type, lag) links: [(12, "FS", 0), (14, "SS", 2)] -> "12, 14SS+2d"."""
    out = []
    for uid, typ, lag in links:
        if not uid:
            continue
        typ, lag = (typ or "FS").upper(), int(lag or 0)
        text = str(int(uid)) + ("" if typ == "FS" and not lag else typ)
        out.append(text + (f"{lag:+d}d" if lag else ""))
    return ", ".join(out)


def take_links(row):
    """Replace a row's P1 / T1 / L1 ... keys (older layouts, the sample, MS Project) with its Predecessors text."""
    links = []
    k = 1
    while any(f"{x}{k}" in row for x in "PTL"):
        links.append((row.pop(f"P{k}", None), row.pop(f"T{k}", None), row.pop(f"L{k}", None)))
        k += 1
    if links and not row.get("PredIn"):
        row["PredIn"] = links_text(links) or None
    return row


def take_people(row, more=()):
    """Replace a row's Res / Units keys, plus more (resource, units) people, with its Resources text:
    "Alex Chen, Jordan Lee[50%]". Used for the sample, MS Project imports and older workbooks."""
    found = [(row.pop("Res", None), row.pop("Units", None))] + list(more)
    text = ", ".join(f"{name}" + ("" if units in (None, "", 1) else f"[{round(units * 100)}%]")
                     for name, units in found if name)
    if text and not row.get("ResIn"):
        row["ResIn"] = text
    return row


def each_link(template):
    """template(k) for every link, comma-separated: the arguments of a MAX over all links."""
    return ",".join(template(k) for k in LINK_NUMBERS)


def over_allocated(resource, start, finish):
    """True when resource is over 100% (Resource Report utilisation, after leave) in a week from start to finish.
    Only the weeks in the Resource Report window and the resources in its grid are known; others count as not over."""
    week_of = lambda d: f"(INT(({d}-ResourceWeekStart)/7)+1)"
    first, last = f"MAX(1,{week_of(start)})", f"MIN(COLUMNS(rw_Util),{week_of(finish)})"
    row = f"MATCH({resource},res_Name,0)"
    return (f"IFERROR(IF({first}>{last},FALSE(),"
            f"COUNTIF(INDEX(rw_Util,{row},{first}):INDEX(rw_Util,{row},{last}),\">1\")>0),FALSE())")


def formulas():
    nm, s, leaf = R("Name"), R("IsSum"), f'{R("IsLeaf")}=1'
    blank = f'{nm}=""'
    manual = f'IF({R("CStart")}>0,{R("CStart")},ProjectStart)'
    f = {}
    # outline
    f["ID"] = f'IF({blank},"",COUNTA({T.upto(N["Name"])})-1)'
    f["Lvl"] = f'IF({blank},"",MAX(1,N({R("Level")})))'
    f["Show"] = f'IF({blank},"",IF({R("Lvl")}<=OutlineShowLevel,"Yes","No"))'
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
    # links: the Predecessors cell (`12, 14SS+2d`, by UID) split into one helper column per link. Read from an input
    # only, so none of this is volatile or recalculated when another cell changes.
    links = f'SUBSTITUTE(SUBSTITUTE(UPPER({R("PredIn")}),";",",")," ","")'
    item_at = lambda k: f'TRIM(MID(SUBSTITUTE({links},",",REPT(" ",99)),{(k - 1) * 99 + 1},99))'
    for k in LINK_NUMBERS:
        item = R(f"I{k}")
        # empty cell, or the previous link was the last: stop before splitting the text again
        before = f'{R("PredIn")}=""' if k == 1 else f'{R(f"I{k - 1}")}=""'
        f[f"I{k}"] = f'IF({before},"",{item_at(k)})'
        cut = f'MIN(FIND("F",{item}&"F"),FIND("S",{item}&"S"),FIND("+",{item}&"+"),FIND("-",{item}&"-"))'
        rest = f"MID({item},{cut},99)"
        typed = f'OR(LEFT({rest},1)="F",LEFT({rest},1)="S")'
        lag = f'SUBSTITUTE(IF({typed},MID({rest},3,99),{rest}),"D","")'
        f[f"P{k}"] = f'IF({item}="","",IFERROR(VALUE(LEFT({item},{cut}-1)),""))'
        f[f"T{k}"] = f'IF({item}="","",IF({typed},LEFT({rest},2),"FS"))'
        f[f"L{k}"] = f'IF({item}="","",IF({lag}="",0,IFERROR(VALUE({lag}),"?")))'
        unreadable = (f'AND({item}<>"",OR({R(f"P{k}")}="",ISNA(MATCH({R(f"T{k}")},{{"FS","SS","FF","SF"}},0)),'
                      f'NOT(ISNUMBER({R(f"L{k}")}))))')
        f[f"R{k}"] = pred_pos(f"P{k}")
        f[f"M{k}"] = (f'IF({unreadable},"can\'t read "&{item}&" (type UIDs like 12 or 12SS+2d, comma-separated)",'
                      f'{link_problem(f"P{k}", f"R{k}")})')
        f[f"D{k}"] = driver(k, "LCS", "LCF")
        f[f"E{k}"] = driver(k, "LFS", "LFF")
    ctype = R("CType")
    f["LinkChk"] = (f'IF({blank},"",TRIM(' + "&".join(
        f'IF({R(f"M{k}")}<>"","Pred {k}: "&{R(f"M{k}")}&". ","")' for k in LINK_NUMBERS) +
        f'&IF(AND({ctype}<>"",N({R("Cons")})=0,{ctype}<>"{CONSTRAINTS[0]}",{ctype}<>"{CONSTRAINTS[1]}"),'
        f'"Constraint ignored: enter a Constraint Date.","")'
        f'&IF({R(f"I{LINKS}")}="","",IF({item_at(LINKS + 1)}<>""," Only the first {LINKS} predecessors are used.",""))))')
    # constraints, as MS Project schedules them (it honours constraint dates over links). The code is MS Project's
    # ConstraintType; a date with a blank Constraint means Start No Earlier Than. _ConstraintStart turns the date into
    # a start date (a finish constraint ends the task on the date); Low / High bound the start: Must Start / Finish On
    # set both, so the task is pinned.
    array = "{" + ",".join(f'"{c}"' for c in CONSTRAINTS) + "}"
    f["CT"] = (f'IF(OR({blank},{s}),"",IF(N({R("Cons")})=0,IF({ctype}="{CONSTRAINTS[1]}",1,0),'
               f'IF({ctype}="",4,IFERROR(MATCH({ctype},{array},0)-1,4))))')
    ct, cd, dur = R("CT"), f'N({R("Cons")})', R("Dur")
    f["CStart"] = (f'IF({cd}=0,0,IF(AND(OR({ct}=3,{ct}=6,{ct}=7),N({dur})>1),'
                   f'WORKDAY.INTL({cd},-(N({dur})-1),{WD}),{cd}))')
    f["CLow"] = f'IF(OR({ct}=2,{ct}=3,{ct}=4,{ct}=6),{R("CStart")},0)'
    f["CHigh"] = f'IF(OR({ct}=2,{ct}=3,{ct}=5,{ct}=7),{R("CStart")},{NO_LIMIT})'
    # detail-task dates (summaries leave these blank)
    f["LCS"] = (f'IF(NOT({leaf}),"",IF({R("Mode")}="Manual",{manual},'
                f'WORKDAY.INTL(MIN(MAX(ProjectStart,{R("CLow")},{each_link(lambda k: R(f"D{k}"))}),{R("CHigh")})-1,1,{WD})))')
    f["LCF"] = f'IF(NOT({leaf}),"",IF(N({R("Dur")})<=0,{R("LCS")},WORKDAY.INTL({R("LCS")},{R("Dur")}-1,{WD})))'
    f["LFS"] = (f'IF(NOT({leaf}),"",IF(N({R("AS")})>0,{R("AS")},IF({R("Mode")}="Manual",{manual},'
                f'WORKDAY.INTL(MAX(MIN(MAX(ProjectStart,{R("CLow")},{each_link(lambda k: R(f"E{k}"))}),{R("CHigh")}),'
                f'IF(RescheduleAfterStatus="Yes",StatusDate,0))-1,1,{WD}))))')
    natural = f'WORKDAY.INTL({R("LFS")},{R("Dur")}-1,{WD})'
    remaining = f'WORKDAY.INTL(StatusWorkday,MAX(0,ROUNDUP({R("Dur")}*(1-MIN(1,N({R("Pct")}))),0)-1),{WD})'
    f["LFF"] = (f'IF(NOT({leaf}),"",IF(N({R("AF")})>0,{R("AF")},IF(N({R("Dur")})<=0,{R("LFS")},'
                f'IF(AND(N({R("AS")})>0,RescheduleAfterStatus="Yes"),MAX({natural},{remaining}),{natural}))))')
    f["LProg"] = f'IF(NOT({leaf}),0,IF(N({R("AF")})>0,1,MIN(1,N({R("Pct")}))))'
    # people: the Resources cell (`Alex Chen, JL[50%]`: names or initials, optional units) split into one set of
    # helper columns per person. The first person uses _Resource1 / _Units1 and the _Leaf... columns; the others
    # their own. Each person's work = Duration x HoursPerDay x their units (MS Project's fixed-units default).
    item = lambda k: f'TRIM(MID(SUBSTITUTE({R("ResIn")},",",REPT(" ",99)),{(k - 1) * 99 + 1},99))'
    person = {1: ("Res", "Units", "LWork", "LCost")} | {k: (f"Rs{k}", f"Ru{k}", f"Rw{k}", f"Rc{k}") for k in MORE_PEOPLE}
    for k, (nm_k, u_k, w_k, c_k) in person.items():
        text = R(f"Ri{k}")                        # this person's part of the cell, split out once
        before = f'OR({blank},{R("ResIn")}="")' if k == 1 else f'{R(f"Ri{k - 1}")}=""'
        f[f"Ri{k}"] = f'IF({before},"",{item(k)})'
        who = f'TRIM(LEFT({text},FIND("[",{text}&"[")-1))'
        f[nm_k] = (f'IF({text}="","",IFERROR(INDEX(res_Name,MATCH({who},res_Name,0)),'
                   f'IFERROR(INDEX(res_Name,MATCH({who},res_Initials,0)),{who})))')
        share = (f'IFERROR(VALUE(SUBSTITUTE(MID({text},FIND("[",{text})+1,FIND("]",{text}&"]")-FIND("[",{text})-1),'
                 f'" ","")),"")')
        f[u_k] = f'IF(ISERROR(FIND("[",{text})),"",IF(N({share})>10,{share}/100,{share}))'
        f[w_k] = (f'IF(OR(NOT({leaf}),{R(nm_k)}=""),0,IF(ISNA(MATCH({R(nm_k)},res_Name,0)),0,'   # unknown: Issues
                  f'N({R("Dur")})*HoursPerDay*IF({R(u_k)}="",DefaultUnits,{R(u_k)})))')
        f[c_k] = f'IF({R(w_k)}=0,0,{R(w_k)}*IFERROR(INDEX(res_Rate,MATCH({R(nm_k)},res_Name,0)),0))'
    # the task's totals add everyone; the per-resource totals (Resources, reports, Sprints) sum each person's columns
    f["NPeople"] = f'IF(NOT({leaf}),0,' + "+".join(f'({R(person[k][0])}<>"")' for k in person) + ")"
    f["LTW"] = "+".join(R(person[k][2]) for k in person)
    f["LTC"] = "+".join(R(person[k][3]) for k in person)
    named = lambda k: (f'IF({R(person[k][0])}="","",", "&{R(person[k][0])}&IF(OR({R(person[k][1])}="",'
                       f'{R(person[k][1])}=1),"","["&ROUND({R(person[k][1])}*100,0)&"%]"))')
    f["ResTxt"] = (f'IF(OR({blank},{s}),"",MID(' + "&".join(named(k) for k in person) +
                   f'&IF({R(f"Ri{PEOPLE}")}="","",IF({item(PEOPLE + 1)}<>""," +more","")),3,999))')
    # critical path: backward pass. Each successor row works out, per link, the latest its predecessor may finish
    # (FS / FF directly; SS / SF via the predecessor's duration, an input) and stores it as a key; the predecessor
    # then takes one MAXIFS over the rows below. A task's late finish is the earliest of those limits, its deadline
    # the project's forecast finish and a Start / Finish No Later Than or Must ... On constraint.
    for k in LINK_NUMBERS:
        valid = f'AND({leaf},{R(f"R{k}")}>0,{R(f"M{k}")}="")'
        typ, lg = f'IF({R(f"T{k}")}="","FS",{R(f"T{k}")})', f"N({R(f'L{k}')})"
        pred_dur = f"N(INDEX({C('Dur')},{R(f'R{k}')}))"
        start_limit = f'WORKDAY.INTL(IF({typ}="SS",{R("LLS")},{R("LLF")}),-{lg},{WD})'
        finish_limit = (f'IF({typ}="FS",WORKDAY.INTL({R("LLS")},-(IF(N({R("Dur")})=0,0,1)+{lg}),{WD}),'
                        f'IF({typ}="FF",WORKDAY.INTL({R("LLF")},-{lg},{WD}),'
                        f'IF({pred_dur}<=0,{start_limit},WORKDAY.INTL({start_limit},{pred_dur}-1,{WD}))))')
        f[f"F{k}"] = f'IF({valid},{NO_LIMIT}-{finish_limit},"")'
        f[f"K{k}"] = f'IF({valid},{SLACK_OFFSET}-{wdiff(R(f"E{k}"), R("LFS"))},"")'
    # where the last successor sits: reads only input columns, so it isn't recalculated after every edit
    f["LastSucc"] = (f'IF(NOT({leaf}),0,MAX('
                     f'{each_link(lambda k: "_xlfn.MAXIFS(" + C("Idx") + "," + C(f"P{k}") + "," + R("UID") + ")")}))')
    succ = f'MAX({each_link(lambda k: successor_max(f"F{k}", f"P{k}"))})'
    f["SuccLF"] = f'IF(NOT({leaf}),"",{NO_LIMIT}-{succ})'
    high = R("CHigh")
    cons_limit = f'IF({high}>={NO_LIMIT},{NO_LIMIT},IF(N({dur})<=1,{high},WORKDAY.INTL({high},{dur}-1,{WD})))'
    f["LLF"] = (f'IF(NOT({leaf}),"",MIN(ProjectFinish,IF(N({R("DL")})>0,{R("DL")},{NO_LIMIT}),{R("SuccLF")},'
                f'{cons_limit}))')
    f["LLS"] = f'IF(NOT({leaf}),"",IF(N({R("Dur")})<=0,{R("LLF")},WORKDAY.INTL({R("LLF")},-({R("Dur")}-1),{WD})))'
    f["LTS"] = f'IF(NOT({leaf}),"",{wdiff(R("LFF"), R("LLF"))})'
    f["LCrit"] = f'AND({leaf},{R("LProg")}<1,N({R("LTS")})<=CriticalSlackDays)'
    slack = f'MAX({each_link(lambda k: successor_max(f"K{k}", f"P{k}"))})'
    f["FSl"] = f'IF(NOT({leaf}),"",IF({slack}=0,{R("LTS")},MIN({R("LTS")},{SLACK_OFFSET}-{slack})))'
    f["LS"] = f'IF({blank},"",IF({s},MIN({block("LLS")}),{R("LLS")}))'
    f["LF"] = f'IF({blank},"",IF({s},MAX({block("LLF")}),{R("LLF")}))'
    f["TS"] = f'IF({blank},"",IF({s},MIN({block("LTS")}),{R("LTS")}))'
    f["Crit"] = (f'IF({blank},"",IF(IF({s},COUNTIF({block("LCrit")},TRUE())>0,{R("LCrit")}),"Yes",""))')
    # every row: detail value, or the roll-up of the detail tasks below a summary
    rollup = lambda fn, leaf_key: f'IF({blank},"",IF({s},{fn}({block(leaf_key)}),{R(leaf_key)}))'
    # As Late As Possible: shown later by the task's free slack, so it finishes just in time for its successors
    # (or the project finish) and moves nothing else. Links, slack and ProjectFinish keep using the early dates
    # (_Leaf...), so this never feeds back into the schedule. The plan uses the same shift as the forecast.
    f["Shift"] = (f'IF(AND({leaf},{R("CT")}=1,{R("Mode")}<>"Manual",N({R("AS")})=0),MAX(0,N({R("FSl")})),0)')
    shown = lambda key: f'IF(NOT({leaf}),"",IF({R("Shift")}=0,{R(key)},WORKDAY.INTL({R(key)},{R("Shift")},{WD})))'
    for key, early in (("XCS", "LCS"), ("XCF", "LCF"), ("XFS", "LFS"), ("XFF", "LFF")):
        f[key] = shown(early)
    f["CPS"] = rollup("MIN", "XCS")
    f["CPF"] = rollup("MAX", "XCF")
    f["FS"] = rollup("MIN", "XFS")
    f["FF"] = rollup("MAX", "XFF")
    f["Work"] = rollup("SUM", "LTW")
    f["Cost"] = rollup("SUM", "LTC")
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
    join = lambda by_id: f'IF({blank},"",MID(' + "&".join(pred_text(k, by_id) for k in LINK_NUMBERS) + ',2,999))'
    f["PredTxt"] = join(True)
    f["UPredTxt"] = join(False)
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
    has_work = f'AND({leaf},{R("LTW")}>0)'
    f["SWD"] = f'IF({has_work},NETWORKDAYS.INTL(ProjectStart,{R("XFS")},{WD}),0)'
    f["FWD"] = f'IF({has_work},MAX({R("SWD")},NETWORKDAYS.INTL(ProjectStart,{R("XFF")},{WD})),0)'
    f["WPD"] = f'IF({has_work},{R("LWork")}/({R("FWD")}-{R("SWD")}+1),0)'
    f["WBefore"] = f'{R("WPD")}*({R("SWD")}-1)'
    for k in MORE_PEOPLE:
        f[f"Rp{k}"] = f'IF(AND({has_work},{R(f"Rw{k}")}>0),{R(f"Rw{k}")}/({R("FWD")}-{R("SWD")}+1),0)'
        f[f"Rb{k}"] = f'{R(f"Rp{k}")}*MAX(0,{R("SWD")}-1)'
    # progress roll-up and the milestone list
    f["WtDur"] = f'IF({leaf},MAX(N({R("Dur")}),0.1),0)'
    f["WtDone"] = f'{R("WtDur")}*{R("LProg")}'
    f["MKey"] = f'IF(AND({leaf},N({R("Dur")})=0,{R("LProg")}<1),{R("XFF")}+{R("Idx")}/100000,"")'
    # earned value at the status date. Budget = Baseline Cost once a baseline is set, otherwise Cost. Planned % =
    # share of the planned working days before the status date (the status date itself is not yet worked).
    # Actual cost = the Actual Cost entered, or Cost x Progress when it is blank (as MS Project calculates it).
    ps, pf = R("PS"), R("PF")
    f["PlanPct"] = (f'IF(NOT({leaf}),0,IF({pf}<StatusDate,1,IF({ps}>=StatusDate,0,'
                    f'IFERROR(MIN(1,NETWORKDAYS.INTL({ps},StatusDate-1,{WD})/NETWORKDAYS.INTL({ps},{pf},{WD})),0))))')
    f["LBud"] = f'IF(NOT({leaf}),0,IF({R("BC")}="",{R("LTC")},N({R("BC")})))'
    f["LPV"] = f'{R("LBud")}*{R("PlanPct")}'
    f["LEV"] = f'{R("LBud")}*{R("LProg")}'
    f["LAC"] = f'IF(NOT({leaf}),0,IF({R("ACin")}="",{R("LTC")}*{R("LProg")},N({R("ACin")})))'
    # the Resource column's person's share of it (by cost; all of it when the task has no cost)
    f["PAC"] = f'IF({R("LTC")}=0,{R("LAC")},{R("LAC")}*{R("LCost")}/{R("LTC")})'
    for k in MORE_PEOPLE:
        f[f"Ra{k}"] = f'IF({R("LTC")}=0,0,{R("LAC")}*{R(f"Rc{k}")}/{R("LTC")})'
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
    f["StartKey"] = (f'IF(AND({leaf},N({R("AS")})=0,{R("XFS")}>=StatusDate,{R("XFS")}<=LookaheadEnd),'
                     f'{R("XFS")}+{tie},"")')
    f["DueKey"] = (f'IF(AND({leaf},{R("LProg")}<1,{R("XFF")}>=StatusDate,{R("XFF")}<=LookaheadEnd),'
                   f'{R("XFF")}+{tie},"")')
    f["CritKey"] = f'IF({R("LCrit")},{R("XFS")}+{tie},"")'
    # resource and agile reports. A task is done on its Actual Finish (or its forecast finish when it is marked
    # 100% without one); story points count on detail tasks only.
    f["LWDone"] = f'{R("LWork")}*{R("LProg")}'
    for k in MORE_PEOPLE:
        f[f"Rd{k}"] = f'{R(f"Rw{k}")}*{R("LProg")}'
    f["DoneDate"] = f'IF(AND({leaf},{R("LProg")}>=1),IF(N({R("AF")})>0,{R("AF")},{R("XFF")}),"")'
    f["LPts"] = f'IF({leaf},N({R("SP")}),0)'
    # over-allocation (Resource Report weekly utilisation, after leave): Yes when any of the task's people is over
    # 100% in a week the task runs. Only weeks in the Resource Report window and its first resources are known.
    # (a sum of IFs, not OR / AND: those would evaluate every person's look-up, including the empty slots)
    over = "+".join(f'IF({R(person[k][0])}="",0,IF({over_allocated(R(person[k][0]), R("XFS"), R("XFF"))},1,0))'
                    for k in person)
    f["OverAlloc"] = f'IF(OR(NOT({leaf}),{R("LProg")}>=1),"",IF({over}>0,"Yes",""))'
    f["CurKey"] = f'IF(AND({leaf},CurrentSprint<>"",{R("Sprint")}=CurrentSprint),{R("Idx")},"")'
    # everything wrong on the row, in words, so a user looks in one place (the check columns keep the detail)
    uid, res = R("UID"), R("Res")
    f["Issues"] = (f'IF({blank},"",TRIM(IF({R("LinkChk")}<>"",{R("LinkChk")}&" ","")'
                   f'&IF(N({uid})=0,"No UID: give the row Settings > NextFreeUID. ",'
                   f'IF(COUNTIF({C("UID")},{uid})>1,"UID "&{uid}&" is used twice. ",""))'
                   f'&IF({R("WBS")}="?","Level is more than 1 below the row above. ","")'
                   + "".join(f'&IF({R(person[k][0])}="","",IF(ISNA(MATCH({R(person[k][0])},res_Name,0)),'
                             f'{R(person[k][0])}&" is not on Resources. ",""))' for k in person)
                   + f'&IF({R(f"Ri{PEOPLE}")}="","",IF({item(PEOPLE + 1)}<>"",'
                     f'"Only the first {PEOPLE} people in Resources are used. ",""))'
                   f'&IF({R("OverAlloc")}="Yes","Over-allocated: see the Resource Report. ","")'
                   f'&IF(LEFT({R("DLChk")},4)="Late","Misses its deadline ("&LOWER({R("DLChk")})&"). ","")))')
    return f


def each(prefix):
    """The per-link keys of one kind: each("P") = P1, P2, ..."""
    return [f"{prefix}{k}" for k in LINK_NUMBERS]


HELPER_DATES = {"CPS", "CPF", "LCS", "LCF", "LFS", "LFF", "XCS", "XCF", "XFS", "XFF", "SuccLF", "LLF", "LLS",
                "DoneDate", "CStart", "CLow", "CHigh", *each("D"), *each("E")}
# key, kind, width, number format, centre, header note
LAYOUT = [
    ("ID", "calc", 5, None, True, "Position in the list; renumbers when rows move (like MS Project)."),
    ("UID", "in", 6, None, True, "Permanent number that predecessors refer to. Each task needs a different one."),
    ("WBS", "calc", 8, None, True, "Outline number, generated from Level."),
    ("Level", "in", 6, None, True, "Outline level: 1 = phase, 2 = task, 3 = sub-task ... (at most 1 more than the row above)."),
    ("Show", "calc", 6, None, True, "Yes when Level <= Settings > OutlineShowLevel: filter on Yes to collapse the "
                                    "outline to that level (or use the 1 2 3 outline buttons at the top left)."),
    ("Name", "in", 36, None, False, None),
    ("Mode", "in", 8, None, True, "Auto = scheduled from predecessors and the constraint. Manual = starts on the "
                                  "Constraint Date (finish constraints: finishes on it)."),
    ("Dur", "in", 9, None, True, "Working days. 0 = milestone."),
    ("PredIn", "in", 16, None, False, "Tasks that must come first, by UID, comma-separated: 12 = after 12 finishes; "
                                      "12SS+2d = 2 days after 12 starts; FS, SS, FF, SF; lag in working days (- = "
                                      f"lead). Up to {LINKS}; each must be a detail task above this one."),
    ("CType", "in", 20, None, False, "MS Project constraint (dropdown). Blank = Start No Earlier Than when a Constraint "
                                     "Date is entered, otherwise As Soon As Possible. Constraint dates win over links."),
    ("Cons", "in", 15, DATE, False, "Date for the Constraint. Manual tasks are pinned to it."),
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
    ("LinkChk", "calc", 24, None, False, "Explains predecessor links and constraints that are ignored."),
    ("Issues", "calc", 34, None, False, "Everything wrong on this row, in words: links, UID, outline level, resource, "
                                        "over-allocation, missed deadline. Blank = fine."),
    ("Pct", "in", 9, "0%", False, "Enter 0-100%. An Actual Finish counts as 100%."),
    ("Prog", "calc", 9, "0%", False, "% Complete, rolled up to summaries (duration-weighted)."),
    ("Status", "calc", 12, None, True, None),
    ("ResIn", "in", 20, None, False, f"Who does it: names or initials from Resources, comma-separated, with an optional "
                                     f"share of their time, e.g. Alex Chen, JL[50%] (blank = DefaultUnits). Up to "
                                     f"{PEOPLE} people."),
    ("ResTxt", "calc", 22, None, False, "Everyone on the task, by full name with units: the Resources cell, checked "
                                        "(MS Project's Resource Names)."),
    ("OverAlloc", "calc", 9, None, True, "Yes when this task's resource is over 100% of capacity (after leave) in a "
                                         "week the task runs: see Resource Report. Only weeks shown there count."),
    ("Work", "calc", 9, "#,##0.0", False, "Hours = Duration x HoursPerDay x Units, for each person on the task."),
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
] + [(k, "help", 11, DATE if k in HELPER_DATES else None, False, None)
     for k in ("Lvl", "Idx", "Parent", "NextLvl", "IsSum", "Desc", "IsLeaf", *each("R"), "CT", "CStart", "CLow",
               "CHigh", "LCS", "LCF", "LFS", "LFF",
               *[f"{x}{k}" for k in LINK_NUMBERS for x in "IPTL"], *[f"Ri{k}" for k in range(1, PEOPLE + 1)], "Res", "Units", "NPeople",
               *[f"{x}{k}" for k in MORE_PEOPLE for x in ("Rs", "Ru", "Rw", "Rc", "Rp", "Rb", "Rd", "Ra")],
               "LProg", "LWork", "LCost", "LTW", "LTC", "PAC", *each("F"), *each("K"), "LastSucc", "SuccLF", "LLF", "LLS",
               "LTS", "LCrit", "Shift", "XCS", "XCF", "XFS", "XFF", "CPS", "CPF", *each("M"), *each("D"), *each("E"),
               "SWD", "FWD", "WPD", "WBefore", "WtDur", "WtDone", "UPredTxt", "MKey",
               "PlanPct", "LBud", "LPV", "LEV", "LAC", "Phase", "PhaseKey", "MsKey", "LateKey", "StartKey", "DueKey", "CritKey",
               "LWDone", "DoneDate", "LPts", "CurKey")]

# The sheet's sections, left to right. The first holds what a plan needs and is always shown; each optional section is
# an Excel column group, collapsed, whose +/- button sits on its first (key) column, which stays visible. (Excel would
# merge groups that touch, so every section keeps one visible column.) Nothing is removed: + shows the rest.
SECTIONS = [
    ("PLAN: type in the blue columns", None, ["ID", "UID", "WBS", "Level", "Name", "Dur", "PredIn",
                                              "ResIn", "Pct", "AS", "AF", "Notes", "FS", "FF", "Status", "Crit", "Issues", "Show"]),
    ("SCHEDULING", "DL", ["Mode", "CType", "Cons"]),
    ("DETAIL", "Var", ["PS", "PF", "SVar", "LS", "LF", "TS", "FSl", "DLChk", "LinkChk", "Prog", "PredTxt"]),
    ("COST", "Work", ["ResTxt", "OverAlloc", "Cost", "ACin"]),
    ("AGILE", "Sprint", ["SprIn", "SP"]),
    ("BASELINE", "DVar",
     ["WVar", "CVar", "SrcS", "SrcF", "SrcD", "SrcW", "SrcC", "BS", "BF", "BD", "BW", "BC"]),
    ("EARNED VALUE", "EV", ["Budget", "PV", "ACt"]),
]
_by_key = {row[0]: row for row in LAYOUT}
_shown = [k for _, key, rest in SECTIONS for k in ([key] if key else []) + rest]
assert len(_shown) == len(set(_shown)) and all(_by_key[k][1] != "help" for k in _shown), "every column once"
LAYOUT = [_by_key[k] for k in _shown] + [row for row in LAYOUT if row[1] == "help"]
assert len(LAYOUT) == len(_by_key), "a visible column is missing from SECTIONS"
STATUS_COLOURS = {"Complete": "548235", "In Progress": "2F75B5", "Behind": "C00000", "Late Start": "C55A11",
                  "Not Started": "7F7F7F"}


def build_tasks(wb, rows, n_rows):
    ws = wb.create_sheet("Tasks")
    title(ws, None)
    ws["A1"] = '=ProjectName&"  -  Task Entry"'
    ws["A2"] = ("Type in BLUE columns only. Level = outline level (1 = phase, 2 = task, 3 = sub-task). Predecessors refer "
                "to the UID column. Fix anything in Issues (red). More columns (scheduling, schedule detail, cost, "
                "agile, baseline, earned value): click + above a section. Insert rows inside the table: formulas fill "
                "in automatically.")
    ws["A2"].font = NOTE_FONT
    ws["A3"] = '="Next free UID for a new row: "&NextFreeUID'
    ws["A3"].font = LABEL_FONT
    f = formulas()
    cols = []
    for key, kind, width, fmt, center, note in LAYOUT:
        values = []
        if key == "UID":
            # stable IDs: links keep working when rows move. Imported rows keep their own; the rest follow on.
            given = [r.get("UID") for r in rows]
            spare = iter(range(max([u for u in given if u] or [0]) + 1, 10 ** 9))
            values = [u if u else next(spare) for u in given] + [next(spare) for _ in range(n_rows - len(rows))]
        elif kind == "in" and any(key in r for r in rows):
            values = [r.get(key) for r in rows]
        cols.append(Col(N[key], kind, width, fmt, None if kind == "in" else f[key], center, note, key, values))
    L = add_table(ws, TABLE, HEADER_ROW, 1, cols, n_rows)
    ws.row_dimensions[HEADER_ROW].height = 42
    ws.sheet_properties.outlinePr.summaryRight = False       # a group's +/- sits on the column left of it
    for label, key, rest in SECTIONS:
        first = L[key] if key else L[rest[0]]
        ws[f"{first}4"] = label
        ws[f"{first}4"].font = LABEL_FONT
        if key:
            group_columns(ws, L[rest[0]], L[rest[-1]], 1)
    ws[f"{L['Lvl']}4"] = "HELPER COLUMNS (hidden)"
    ws[f"{L['Lvl']}4"].font = LABEL_FONT
    for key, defined in NAMES.items():
        name(wb, defined, C(key))
    ws.freeze_panes = f"{L['Dur']}{HEADER_ROW + 1}"
    group_outline(ws, [r.get("Level") for r in rows], HEADER_ROW + 1)
    for i, r in enumerate(rows):                              # names indented by level, for the rows built in
        level = r.get("Level")
        if level and level > 1:
            ws[f"{L['Name']}{HEADER_ROW + 1 + i}"].alignment = Alignment(indent=min(int(level) - 1, 15))
    group_columns(ws, L["Lvl"], L["CurKey"], 2)
    _validation_and_formatting(ws, L, n_rows)
    ws.sheet_view.zoomScale = 90
    return ws


def group_columns(ws, first, last, level):
    """Collapse columns first..last (letters) into an Excel outline group. Each column keeps its own width (openpyxl's
    column_dimensions.group merges them into one and loses the widths)."""
    for c in range(column_index_from_string(first), column_index_from_string(last) + 1):
        dim = ws.column_dimensions[get_column_letter(c)]
        dim.outlineLevel, dim.hidden = level, True
    ws.sheet_format.outlineLevelCol = max(ws.sheet_format.outlineLevelCol or 0, level)


def group_outline(ws, levels, first_row):
    """Excel row groups by outline level (summary above its tasks), so the 1 2 3 ... buttons at the top left expand
    and collapse the plan as in MS Project. Grouping is part of the sheet, not a formula: it is set for the rows the
    workbook is built with; rows added later are grouped with Data > Group (or use the Show filter)."""
    ws.sheet_properties.outlinePr.summaryBelow = False
    deepest = 0
    for i, level in enumerate(levels):
        if level and level > 1:
            ws.row_dimensions[first_row + i].outline_level = depth = min(int(level) - 1, 7)
            deepest = max(deepest, depth)
    ws.sheet_format.outlineLevelRow = deepest          # Excel shows that many outline buttons


# What to type, shown when a cell in the column is selected (Excel's input message: desktop, web and iPhone).
# Excel limits the title to 32 characters and the text to 255.
PROMPTS = {
    "UID": ("UID", "Pre-filled: a permanent number that other tasks' predecessors refer to. A new row: "
                   "Settings > NextFreeUID."),
    "Level": ("Outline level", "1 = phase, 2 = task, 3 = sub-task. At most 1 more than the row above. A row with deeper "
                               "rows below it becomes a summary."),
    "Name": ("Task name", "What the task is, e.g. Build login page. Leave the row blank to end the plan."),
    "Dur": ("Duration (working days)", "e.g. 5. 0 = a milestone. Summaries roll up: leave them blank."),
    "PredIn": ("Predecessors (by UID)", "e.g. 12, 14SS+2d: after 12 finishes and 2 working days after 14 starts. "
                                        "Types FS (default), SS, FF, SF; - = lead. Each must be a detail task above "
                                        "this one."),
    "ResIn": ("Resources", "Who does it: names or initials from Resources, comma-separated. Add a share of "
                           "their time in brackets, e.g. Alex Chen, JL[50%]. Pick one from the list or type several."),
    "Pct": ("% Complete", "0% - 100%. An Actual Finish counts as 100%."),
    "AS": ("Actual Start", "The day work really started. Leave blank until it starts."),
    "AF": ("Actual Finish", "The day it really finished: the task counts as 100% complete."),
    "Notes": ("Notes", "Anything worth knowing: assumptions, links, who to ask."),
    "DL": ("Deadline", "Optional: the date it must be done by. Issues and the Gantt flag a forecast after it."),
}


def _validation_and_formatting(ws, L, n_rows):
    """Data validation and conditional formatting: Excel only accepts A1 ranges here (it shifts them on insert)."""
    first, last = HEADER_ROW + 1, HEADER_ROW + n_rows
    rng = lambda k: col_range(L[k], first, last)
    at = lambda k: f"${L[k]}{first}"           # this row's cell, relative to the top-left of the formatted range

    def dv(rule, *keys):
        """One rule per column, so each can carry its own input prompt (PROMPTS)."""
        for k in keys:
            r = copy.copy(rule)
            r.sqref = MultiCellRange()
            if k in PROMPTS:
                r.showInputMessage, (r.promptTitle, r.prompt) = True, PROMPTS[k]
            ws.add_data_validation(r)
            r.add(rng(k))

    dv(DataValidation(type="whole", operator="greaterThanOrEqual", formula1="1", showErrorMessage=True, errorTitle="UID",
                      error="UID must be a whole number; give each task a different one (see Settings > NextFreeUID)."), "UID")
    dv(DataValidation(type="whole", operator="between", formula1="1", formula2="9", showErrorMessage=True,
                      errorTitle="Level", error="Outline level must be 1-9 and at most 1 more than the row above."), "Level")
    dv(DataValidation(type="list", formula1="=L_Mode"), "Mode")
    dv(DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", showErrorMessage=True,
                      error="Duration is in working days (0 = milestone)."), "Dur")
    dv(DataValidation(type="list", formula1="=L_Constraint"), "CType")
    dv(DataValidation(type="date", operator="greaterThan", formula1="1"), "Cons", "DL", "AS", "AF")
    dv(DataValidation(type="decimal", operator="between", formula1="0", formula2="1", showErrorMessage=True,
                      error="Enter 0% - 100%."), "Pct")
    dv(DataValidation(type="list", formula1="=res_Name", showErrorMessage=False), "ResIn")   # pick one, or type several
    dv(DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0"), "ACin", "SP")
    dv(DataValidation(type="list", formula1="=spr_Name"), "SprIn")
    dv(DataValidation(allow_blank=True), "Name", "Notes", "PredIn")          # prompts only: any text

    cf = ws.conditional_formatting
    cf.add(f"A{first}:{L['ACt']}{last}", FormulaRule(formula=[f"{at('IsSum')}=TRUE"], font=Font(bold=True), fill=fill("DDEBF7")))
    for k in ("Dur", "PredIn", "CType", "Cons", "AS", "AF", "Pct", "ResIn", "ACin", "SprIn", "SP"):
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
