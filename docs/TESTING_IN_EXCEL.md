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
      1.0" and page numbers.
- [ ] Settings > TakeBaseline = Yes, then TakeSnapshot's two baseline links (copy, Paste Special > Values): the Gantt
      outlines follow the current dates. Set it back to No.
- [ ] Tick each line of [RELEASE_CHECKLIST.md](RELEASE_CHECKLIST.md) marked manual.
