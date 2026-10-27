# Testing the planner on an iPhone

Covers backlog item B-001 until desktop Excel is available.

## What an iPhone can and can't prove

| Check | iPhone? | How |
|---|---|---|
| Formulas calculate in real Excel (no `#NAME?`, no circular-reference warning) | ✅ | Microsoft Excel app: same calculation engine as desktop Excel |
| Conditional formatting (Gantt bars), dropdowns, frozen panes, charts | ✅ | Excel app |
| Excel for the web | ✅ | Safari → office.com → upload to OneDrive → open |
| Macros (.xlsm) | ❌ | Excel on iPhone, iPad and the web cannot run VBA |
| Import into Microsoft Project | ❌ | MS Project desktop is Windows-only; needs a Windows PC (or a cloud Windows VM) with Project installed |
| Apple Numbers | ⚠️ | Opens the file but uses a different engine; results don't prove Excel compatibility. Skip it. |

## Setup (about 5 minutes)

1. Install **Microsoft Excel** from the App Store. Sign in with a free Microsoft account; editing is free on phones.
2. Get the file onto the phone, either way:
   - In the Claude app, open the `Project_Planner.xlsx` card → Share → **Excel**; or
   - Safari → `github.com/LoneWolfDen/excel-project-planner` (sign in, it's private) → `Project_Planner.xlsx` → **View raw** → Share → **Excel**.
3. If Excel offers **Enable editing**, tap it. Rotate the phone to landscape for the Tasks and Gantt sheets.

## Checklist

Tick each one and note anything odd (a screenshot is ideal).

**Opening**
- [ ] Opens with no "repair" or "we found a problem with some content" message.
- [ ] No circular-reference warning, on open or after edits below.

**Calculation (Tasks sheet)**
- [ ] No `#NAME?`, `#VALUE!` or `#REF!` anywhere. Pay attention to the Planned Start / Finish columns.
- [ ] Task 19 "Build complete": Deadline Check shows "Late by …" in red; Link Check column is empty for every task.
- [ ] Task 2 "Project charter": Planned Start Mon 07-Sep-26, Planned Finish Wed 09-Sep-26.
- [ ] Task 7 "Solution architecture" (SS + 3): Planned Start Thu 17-Sep-26.
- [ ] Task 26 "Hypercare support": Planned Finish Mon 04-Jan-27 (holidays skipped).
- [ ] Task 13 "Core API endpoints" status "Behind" (the sample's status date is 25-Oct-2026, on Settings).

**Ripple test (edit something)**
- [ ] Change Task 2's Duration from 3 to 5. Planned dates of tasks 2, 3, 4 and later move by 2 working days.
- [ ] Undo, or set it back to 3.
- [ ] Type `7` in Task 7's "Predecessors" (a link to itself). Link Check shows "Pred 1: links to itself." and no circular-reference warning appears. Undo.
- [ ] Type `6, 21` in Task 8's "Predecessors" (a second link, to a task further down). Link Check shows "Pred 2: UID 21 is below this task: move it above." Undo.
- [ ] On Settings, change `GanttPeriod` from 1 to 7. The Gantt switches to weeks. Set it back to 1.

**Gantt sheet**
- [ ] Green / blue / orange / dark grey bars appear; planned dates show as a thin dark outline; milestones are purple cells.
- [ ] Red lines at the status date (25-Oct), grey weekend columns, red deadline cells on tasks 19 and 25.

**Baseline**
- [ ] On Tasks, select the 5 "BL Source" columns for rows 6–35, Copy, tap the first "Baseline Start" cell, Paste Special → Values. Status > KEY FIGURES "Baseline" shows "Set on 30 tasks".
- [ ] Change Task 21's Duration from 8 to 10: Planned Finish stays put, Finish moves 2 days, Finish Variance shows 2.

**Inserting a task**
- [ ] On Tasks, long-press row 19's number → Insert (row above). The new row fills in its grey formula cells by itself, and its UID cell turns red until you type the number shown at the top of the sheet ("Next free UID").

**Other sheets**
- [ ] Status: change the status date at the top; the key figures, RAG and charts update, and Print preview fits one page.
- [ ] Log: pick Type = Decision on an empty row; the RAID columns turn grey and the Status dropdown offers Proposed, Approved ...
- [ ] Agile: the sprint table, release figures and charts show. Resource Report: hours and utilisation by sprint fill in; red cells only where someone is over capacity.
- [ ] Tasks: the Resource and Type cells show dropdowns.

## Reporting back

Paste the results (or screenshots) into the B-001 GitHub issue or this chat. A failure in "Calculation" is the most important to report, with the cell and the value shown.
