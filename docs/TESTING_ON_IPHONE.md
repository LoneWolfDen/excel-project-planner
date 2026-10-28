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
3. If Excel offers **Enable editing**, tap it. Rotate the phone to landscape for the Tasks sheet.

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

**Gantt timeline** (Tasks, right of the table: tap "Gantt timeline >>" at the top)
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
- [ ] Agile: the sprint table, release figures and charts show; picking a sprint under "Sprint to show" lists its tasks. Resources: the person-days and points grids fill in. Cost: task costs and earned value show; deleting the sheet leaves no #REF! elsewhere. Resource Report: hours and utilisation by sprint fill in; red cells only where someone is over capacity.
- [ ] Tasks: the Resource and Type cells show dropdowns.

## Release 1.0 (`release-1.0/Project_Planner.xlsx`, about 10 minutes)

Open the Release 1.0 sample the same way. On a phone the dropdowns open as a list at the bottom of the screen.

- [ ] Opens with no repair message; no `#` errors (the sample's status date is 25-Oct-2026).
- [ ] Tasks: Pred1 on task 14 shows a dropdown of "UID Name" entries; pick "12 Data model & database setup": Start
      moves, no Issues. Undo. The input message shows when the cell is selected.
- [ ] Type `2days` in a Duration: refused. `2w` is taken; the row's Finish moves to 10 working days. Undo.
- [ ] Type a UID that another task has (e.g. `5` in task 6's UID): refused.
- [ ] Res1 offers the people on Resources; Units1 takes 50%.
- [ ] Resources > LEAVE BY DAY (under the people): Casey Kim's November row has 1s and a 0.5; its Days column reads
      "16, 17(0.5), 18 ...". Type 1 under a weekday for Sam Rivera in a new row (pick the person and month): the
      Resource Report's leave for that week rises.
- [ ] Check: the counts, "n to fix before you report" in red, and the lists of tasks with Issues, over-allocated tasks
      and overdue Log items. Fix one (e.g. close the overdue Log item): the count drops.
- [ ] TakeSnapshot: each link jumps to its place (the snapshot row, History, the BL Source and Baseline columns).
- [ ] Status, Check, Log: typing in a grey cell is refused; the blue cells take it. Tasks: Insert a table row works.
- [ ] Simple (`release-1.0/Project_Planner_Simple.xlsx`): six sheets, 12 tasks, no errors.

## Timing in Microsoft 365 (desktop, about 5 minutes)

CI times recalculation in LibreOffice only. Once per release, time the same edits in Excel (B-1016):

1. Build the 1,000-task workbook: `python scripts/build_planner.py stress.xlsx --tasks 1000 --stress 33`, and open
   it in Excel for Microsoft 365 on a desktop.
2. Formulas > Calculation Options > Manual. Then Ctrl+Alt+F9 (full recalculation) and time it with a stopwatch;
   repeat twice and keep the middle time.
3. Change the Duration of task 2, then press F9; time that too (middle of three).
4. Switch back to Automatic and report both times with the Excel build (File > Account > About Excel) next to the
   CI numbers in the release PR. Excel is usually several times faster than LibreOffice; a slower one is a bug.

## Reporting back

Paste the results (or screenshots) into the B-001 GitHub issue or this chat. A failure in "Calculation" is the most important to report, with the cell and the value shown.
