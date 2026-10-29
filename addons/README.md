# Optional add-ons (B-704, B-705)

The planner is a plain `.xlsx` and needs none of this: every action below is a copy-and-paste or a typed value the
TakeSnapshot sheet and Start > Inserting tasks describe. The add-ons only save clicks. Everything is found by
defined name or table name (`tsk_BLSource`, `snap_Row`, the `Tasks` table, ...), never by cell address, so inserted
rows and columns don't break them.

| Action | By hand | VBA macro (`vba/PlannerMacros.bas`) | Office Script (`office-scripts/`) |
|---|---|---|---|
| Set baseline | TakeSnapshot > BASELINE: copy BL Source, paste values on Baseline | `SetBaseline` | `setBaseline.ts` |
| Keep a baseline as 2 / 3 (template 1.2) | TakeSnapshot > KEEP A BASELINE | `KeepBaseline` | `keepBaseline.ts` |
| Record snapshot | TakeSnapshot > SNAPSHOT: copy the snapshot row, paste values on History | `RecordSnapshot` | `recordSnapshot.ts` |
| Insert task | Start > Inserting tasks: insert a row, type NextFreeUID | `InsertTask` | `insertTask.ts` |
| Fill missing UIDs | Check > Named tasks without a UID: type NextFreeUID into each | `FillMissingUIDs` | `fillUids.ts` |
| Indent / outdent | change Level | `IndentTask`, `OutdentTask` | `indentTask.ts` (`by` = 1 or -1) |
| Expand / collapse | the outline buttons 1 / 9 | `ExpandAll`, `CollapseAll` | (Excel for the web's outline buttons) |

## VBA (desktop Excel)
1. Save a copy of the planner as **Excel Macro-Enabled Workbook (.xlsm)**. Keep the `.xlsx` for sharing: Excel for
   the web, iPhone and iPad can't run macros.
2. Developer > Visual Basic > File > Import File > `vba/PlannerMacros.bas`.
3. Add buttons (Developer > Insert > Button) or Quick Access Toolbar entries for the macros you want.

## Office Scripts (Excel for the web, Microsoft 365 business accounts)
Automate > New Script, paste a `.ts` file, Save; run it from the Automate tab or add it as a button
(… > Add in workbook). Scripts with a parameter (`which`, `by`) ask for it when run from Power Automate; from the
Automate tab they use the default.

These are checked by structure in `tests/test_addons.py` (each names only things the workbooks define) and by hand in
Excel (docs/TESTING_IN_EXCEL.md).
