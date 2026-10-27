"""MS Project export sheets. Column headers are MS Project field names, so the Import Wizard maps them itself.
Rates and costs aren't here: they live on the Cost sheet, which nothing else reads (the XML converter carries them)."""
from .core import SHORT_DATE, Col, Tab, add_table
from .tasks import CONSTRAINTS


def build_msp_tasks(wb, n_rows):
    ws = wb.create_sheet("MSP_Tasks")
    m = Tab("MSPTasks")
    n = f"(ROW()-ROW({m.head('ID')}))"
    # Rows past the end of the Tasks table (after tasks are deleted) are blank. Every formula tests that first with
    # a plain IF, never inside OR(), because OR evaluates all of its arguments and INDEX past the end is #REF!.
    blank = f'IFERROR(INDEX(tsk_Name,{n})="",TRUE())'
    t = lambda col: f"INDEX({col},{n})"
    show = lambda col: f'IF({blank},"",{t(col)})'
    not_summary = lambda expr: f'IF({blank},"",IF({t("tsk_IsSummary")},"",{expr}))'
    date = lambda col: f'IF({blank},"",IF(N({t(col)})=0,"",{t(col)}))'
    code, names = t("tsk_ConstraintCode"), ",".join(f'"{c}"' for c in CONSTRAINTS)
    cols = [
        Col("ID", width=5, formula=show("tsk_ID")),
        Col("Unique ID", width=8, formula=show("tsk_UID")),
        Col("Name", width=36, formula=show("tsk_Name")),
        Col("Outline Level", width=8, formula=show("tsk_Level")),
        Col("Task Mode", width=18, formula=f'IF({blank},"",IF({t("tsk_Mode")}="Manual","Manually Scheduled","Auto Scheduled"))'),
        Col("Duration", width=9, formula=not_summary(f'N({t("tsk_Duration")})&"d"')),
        Col("Start", width=12, fmt=SHORT_DATE, formula=show("tsk_FcStart")),
        Col("Finish", width=12, fmt=SHORT_DATE, formula=show("tsk_FcFinish")),
        Col("Predecessors", width=14, formula=not_summary(t("tsk_Predecessors"))),
        Col("Unique ID Predecessors", width=14, formula=not_summary(t("tsk_UIDPredecessors"))),
        Col("Resource Names", width=22, formula=f'IF({blank},"",{t("tsk_ResourceNames")}&"")'),
        Col("% Complete", width=9, formula=not_summary(f'ROUND({t("tsk_Progress")}*100,0)')),
        Col("Actual Start", width=12, fmt=SHORT_DATE, formula=date("tsk_ActStart")),
        Col("Actual Finish", width=12, fmt=SHORT_DATE, formula=date("tsk_ActFinish")),
        Col("Baseline Start", width=12, fmt=SHORT_DATE, formula=show("tsk_PlanStart")),
        Col("Baseline Finish", width=12, fmt=SHORT_DATE, formula=show("tsk_PlanFinish")),
        Col("Baseline Duration", width=9, formula=f'IF({blank},"",IF({t("tsk_BaseDuration")}="","",N({t("tsk_BaseDuration")})&"d"))'),
        Col("Baseline Work", width=9, formula=f'IF({blank},"",IF({t("tsk_BaseWork")}="","",N({t("tsk_BaseWork")})&"h"))'),
        Col("Constraint Type", width=22, formula=not_summary(
            f'IF({t("tsk_Mode")}="Manual","{CONSTRAINTS[0]}",CHOOSE({code}+1,{names}))')),
        Col("Constraint Date", width=12, fmt=SHORT_DATE, formula=(
            f'IF({blank},"",IF(OR({t("tsk_IsSummary")},{t("tsk_Mode")}="Manual"),"",IF(N({code})<2,"",'
            f'{t("tsk_Constraint")})))')),
        Col("Deadline", width=12, fmt=SHORT_DATE, formula=date("tsk_Deadline")),
        Col("Text1", width=10, formula=not_summary(f'{t("tsk_Sprint")}&""'),
            note="Sprint, in MS Project's custom field Text1 (rename it Sprint in MS Project)."),
        Col("Number1", width=8, formula=not_summary(f'IF(N({t("tsk_StoryPoints")})=0,"",{t("tsk_StoryPoints")})'),
            note="Story points, in MS Project's custom field Number1."),
        Col("Notes", width=40, formula=f'IF({blank},"",{t("tsk_Notes")}&"")'),
    ]
    add_table(ws, "MSPTasks", 1, 1, cols, n_rows)
    ws.freeze_panes = "C2"
    return ws


def build_msp_resources(wb, n_rows):
    ws = wb.create_sheet("MSP_Resources")
    m = Tab("MSPResources")
    n = f"(ROW()-ROW({m.head('Name')}))"
    blank = f'IFERROR(INDEX(res_Name,{n})="",TRUE())'
    t = lambda col: f"INDEX({col},{n})"
    text = lambda col: f'IF({blank},"",{t(col)}&"")'
    cols = [
        Col("Name", width=22, formula=text("res_Name")),
        Col("Initials", width=9, formula=text("res_Initials")),
        Col("Group", width=20, formula=text("res_Group")),
        Col("Type", width=10, formula=text("res_Type")),
        Col("Max Units", width=10, formula=f'IF({blank},"",ROUND({t("res_Allocation")}*100,0)&"%")'),
        Col("Email Address", width=26, formula=text("res_Email")),
    ]
    add_table(ws, "MSPResources", 1, 1, cols, n_rows)
    return ws
