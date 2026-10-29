# Testing the planner in Excel desktop

About 30 minutes, with Excel 2019 or Microsoft 365 on Windows or Mac. Open the committed `Project_Planner.xlsx`
(the sample, with a baseline). Keep `Project_Planner_Simple.xlsx` and `Project_Planner_Blank.xlsx` for the last
section. If Excel asks, click **Enable Editing**. Report a failure with the sheet, the cell and what it showed.

## 1. It calculates

- [ ] Formulas > Calculation Options shows **Automatic**.
- [ ] No `#REF!`, `#NAME?` or `#VALUE!` on any sheet (Home > Find > `#`, Look in: Values).
- [ ] Tasks: change task 13's Stage to Complete. Its Finish, the tasks after it, Status and the Roadmap all change
      at once. Undo (Ctrl+Z).

## 2. One plan sheet (Tasks)

- [ ] Always visible: ID, UID, Level, Task Name, Duration, Predecessors, Resources, Stage, Actual Start / Finish,
      Start, Finish, Sprint, Notes, Issues. The + buttons above row 4 open Scheduling, Detail, Work, Agile and
      Baseline.
- [ ] Click **Gantt timeline >>** (row 3): the timeline sits to the right on the same rows.
- [ ] Black outline = planned (baseline) dates; fill = current schedule (green done, blue remaining, orange
      critical, red where it runs past the planned finish). Task 13 shows a red tail.
- [ ] Insert a row inside the table above task 14 (right-click > Insert > Table Rows Above), type a name, Duration 2,
      Predecessors `13`: every formula fills in and its bar appears.
- [ ] Paste `2d` into a Duration and `18.09.2026` (as text) into an Actual Start: Issues says each isn't a number /
      date. Type `Casey` in Resources: "Did you mean Casey Kim?". Undo.

## 3. Sprints and capacity

- [ ] Agile: type a **Set End** on sprint 4; sprint 5 onwards move after it. Undo.
- [ ] Resources: change Jordan Lee's Allocation % to 50%. The sprint table's Available Days and Capacity drop.
      Undo.
- [ ] Agile > Sprint to show: pick Sprint 5. Its tasks, points and burndown show.

## 4. Reports

- [ ] Status: RAG by area, key figures, the five late / due / starting / critical tasks. Print preview fits one page.
- [ ] Roadmap: each phase has an owner, progress with "n of m done", a RAG, its next milestone, and **Go to phase**
      jumps to its row.
- [ ] Status, Roadmap and Resource Report are protected (typing in a grey cell is refused); the blue cells (PM RAG,
      Comment, notes, status date) still take typing. Review > Unprotect Sheet works without a password.
- [ ] Settings > Live resource grids = No: the Resource Report grids go blank and Status says so in red. Set Yes.

## 5. Cost (optional sheet)

- [ ] Cost: change Priya Sharma's Day Rate. Task cost, actual cost and earned value change.
- [ ] Delete the Cost sheet (right-click the tab > Delete). No `#REF!` anywhere else. Close without saving.

## 6. The other workbooks and the converter page

- [ ] `Project_Planner_Simple.xlsx`: six sheets (Start, Status, Roadmap, Settings, Tasks, Resources), no errors,
      the plan schedules the same as the full sample.
- [ ] `Project_Planner_Blank.xlsx`: on Start pick Project type **Construction**. Settings > Stages shows Mobilised,
      Under Way ...; sprint length 28.
- [ ] Open `docs/export.html` in a browser and try each: workbook to MS Project XML (the notes list what MS Project
      won't see); paste an outline ("Design / two indented tasks; 3d; after 2") and download a workbook; Reorder
      tasks on a workbook where Link Check says a predecessor is below.
- [ ] Optional, for timings: the "Timing in Microsoft 365" steps in [TESTING_ON_IPHONE.md](TESTING_ON_IPHONE.md).

## 7. Release 1.0 (`release-1.0/`)

- [ ] The "Release 1.0" steps in [TESTING_ON_IPHONE.md](TESTING_ON_IPHONE.md), on the desktop.
- [ ] Review > Protect Sheet on Tasks (no password): typing in Finish is refused, Duration and Pred1 still take
      typing. Unprotect, then insert a table row: formulas fill in.
- [ ] File > Print preview on Status, Roadmap and Resource Report: one page each, footer "Project Planner template
      1.0" and page numbers. Page Layout > Print Titles shows the title rows (project name, status date).
- [ ] Tasks > Pred1 on Build complete (UID 19): the dropdown lists only the tasks above it, and opens at once (also
      in Excel for the web). On the first task it lists one entry.
- [ ] Resources: type 30 in Jordan Lee's Oct-26 cell (30-Oct is already on the LEAVE BY DAY grid): the grid's Also
      Typed Above shows 30 in red and Check counts "Leave typed twice". Undo.
- [ ] Settings > TakeBaseline = Yes, then TakeSnapshot's two baseline links (copy, Paste Special > Values): the Gantt
      outlines follow the current dates. Set it back to No.
- [ ] Tick each line of [RELEASE_CHECKLIST.md](RELEASE_CHECKLIST.md) marked manual.

## 8. Microsoft 365 edition (`release-1.1-m365/`)

Excel for Microsoft 365 (or Excel 2021 or later) only. The tests can't calculate this edition (LibreOffice has no
LAMBDA or FILTER), so this section is its real test. The file opens without values: give Excel a moment to
calculate. Every value should equal the 1.1 sample's (`release-1.1/Project_Planner.xlsx`).

- [ ] Opens without a repair message; no `#NAME?` anywhere (Home > Find > `#`, Look in: Values).
- [ ] Formulas > Name Manager lists `L_PARSE_DURATION`, `L_WORKING_DAYS` and `L_IS_WORKING_DAY`, each starting
      `=LAMBDA(`.
- [ ] Tasks: Hypercare support (UID 26, Duration `2w`) has `_Days` 10; Training (UID 23, `1w`) has 5. Type `3d` in a
      Duration: its `_Days` is 3. Undo.
- [ ] Settings: Project finish variance 2; Finish P50 Fri 08-Jan-2027; Finish P80 Tue 12-Jan-2027; forecast finish
      Wed 06-Jan-2027.
- [ ] Log: Dependency 13's Gap (days) is -2; Dependency 5's is 1.
- [ ] Check: "9 to fix before you report"; counts 4 (Issues), 0, 0, 0, 3 (over-allocated), 1 (overdue), 0, 0, 1
      (late dependency). The lists show tasks 13, 19, 23, 26 with issues; 13, 23, 26 over-allocated; Log item 3
      overdue. Each list is one formula that spills (a blue border when a result is selected; no curly braces in
      the formula bar), not `#SPILL!`; the rows below it are blank, not `#N/A`.
- [ ] Tasks > Pred1 on Go-live (UID 25): the dropdown lists only the tasks above it.
- [ ] Type a Saturday in an Actual Start: Issues says "Actual Start is not a working day". Undo.
- [ ] Settings > Sandbox = Yes: the red banner on Status, Core API endpoints (UID 13) gets 7 days. Set No.
- [ ] Save, then export it with docs/export.html (the converter reads the values Excel saved).

## 9. Optional add-ons (`addons/`)

Only for teams that use them (addons/README.md). In a copy saved as `.xlsm` with `vba/PlannerMacros.bas` imported,
or in Excel for the web with the scripts in `office-scripts/`, on the template 1.2 sample:

- [ ] SetBaseline / setBaseline.ts: the Baseline columns equal BL Source (as TakeSnapshot's steps give); TakeBaseline
      is No.
- [ ] KeepBaseline / keepBaseline.ts (2): Baseline 2 Start / Finish equal Baseline Start / Finish.
- [ ] RecordSnapshot / recordSnapshot.ts: one new History row, values not formulas, equal to the snapshot row.
- [ ] InsertTask / insertTask.ts on Core API endpoints: a row below it at level 3 with UID = the old NextFreeUID; the
      formula columns filled; no `#REF!`.
- [ ] IndentTask / OutdentTask / indentTask.ts: Level changes by one; the WBS and the outline follow.
