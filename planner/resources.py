"""Resource sheet (Excel Table "ResourceTable")."""
from openpyxl.worksheet.datavalidation import DataValidation

from .core import Col, Tab, add_table, col_range, name, title

RESOURCES = [("Priya Sharma", "PS", "Project Manager", 0.5, 95, "priya@example.com"),
             ("Morgan Patel", "MP", "Business Analyst", 1, 70, "morgan@example.com"),
             ("Alex Chen", "AC", "Solution Architect", 0.5, 110, "alex@example.com"),
             ("Casey Kim", "CK", "UX Designer", 1, 75, "casey@example.com"),
             ("Jordan Lee", "JL", "Backend Developer", 1, 85, "jordan@example.com"),
             ("Sam Rivera", "SR", "Frontend Developer", 1, 80, "sam@example.com"),
             ("Taylor Morgan", "TM", "QA Engineer", 1, 65, "taylor@example.com")]

NAMES = {"Resource Name": "res_Name", "Initials": "res_Initials", "Group": "res_Group", "Type": "res_Type",
         "Max Units": "res_MaxUnits", "Std Rate": "res_Rate", "Email": "res_Email"}


def build_resources(wb, blank, n_rows=50):
    ws = wb.create_sheet("Resources")
    title(ws, "Resource Sheet", "Like MS Project's Resource Sheet. Names here feed the Resource dropdown on Tasks. "
                                "Add rows to the table for more people.")
    ws.column_dimensions["A"].width = 3
    data = [] if blank else RESOURCES
    t = Tab("ResourceTable")
    nm = t.row("Resource Name")
    letters = add_table(ws, "ResourceTable", 5, 2, [
        Col("Resource Name", "in", 22, values=[d[0] for d in data]),
        Col("Initials", "in", 9, values=[d[1] for d in data]),
        Col("Group", "in", 20, values=[d[2] for d in data]),
        Col("Type", "in", 10, values=["Work"] * len(data), note="Work, Material or Cost (MS Project resource types)."),
        Col("Max Units", "in", 10, "0%", values=[d[3] for d in data], note="Availability: 100% = full time."),
        Col("Std Rate", "in", 12, "#,##0.00", values=[d[4] for d in data], note="Standard rate per hour."),
        Col("Email", "in", 26, values=[d[5] for d in data]),
        Col("Tasks", "calc", 9, "0", f'IF({nm}="","",COUNTIFS(tsk_Resource,{nm},tsk_IsLeaf,1))',
            note="Detail tasks assigned to this resource."),
        Col("Work", "calc", 13, "#,##0.0", f'IF({nm}="","",SUMIFS(tsk_Work,tsk_Resource,{nm},tsk_IsLeaf,1))',
            note="Assigned work (hours)."),
        Col("Cost", "calc", 13, "#,##0.00", f'IF({nm}="","",SUMIFS(tsk_Cost,tsk_Resource,{nm},tsk_IsLeaf,1))',
            note="Assigned cost."),
        Col("Actual Cost", "calc", 13, "#,##0.00", f'IF({nm}="","",SUMIFS(tsk_LeafAC,tsk_Resource,{nm}))',
            note="Cost to date: Actual Cost entered on Tasks, or Cost x Progress where blank."),
        Col("Notes", "in", 30),
    ], n_rows)
    for column, defined in NAMES.items():
        name(wb, defined, f"ResourceTable[{column}]")
    dv = DataValidation(type="list", formula1="=L_ResType")
    ws.add_data_validation(dv)
    dv.add(col_range(letters["Type"], 6, 5 + n_rows))
    ws.freeze_panes = "C6"
    return ws
