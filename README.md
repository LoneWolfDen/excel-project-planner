# Excel Project Planner (MS Project-compatible)

An Excel workbook that works like Microsoft Project: tasks and sub-tasks, dependencies,
planned / actual / forecast dates, a Gantt chart view, resources and sprint capacity.
Built on Excel Tables: formulas use structured references (`[@Duration]`) and named ranges (`ProjectStart`,
`tsk_PlanStart`, `res_Rate`), never A1 cell addresses, so inserting rows or columns never breaks them. Inserted
table rows get their formulas automatically. Recalculation cost is measured on every PR. See
[docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).

**Requires Excel 2019 or Microsoft 365** (Windows or Mac; Excel for the web works too).

## Files
- `Project_Planner.xlsx` – the template, loaded with a 30-task sample project.
- `Project_Planner_Blank.xlsx` – the same template with no sample data, for a new project.
- `docs/` – [architecture and rules](docs/ARCHITECTURE.md), [scheduling spec](docs/SCHEDULING_SPEC.md), [MS Project gap analysis](docs/MSP_GAP_ANALYSIS.md), [backlog](docs/BACKLOG.md), [testing on iPhone](docs/TESTING_ON_IPHONE.md).
- `planner/` – the generator: one module per sheet (`tasks.py`, `gantt.py`, `agile.py`, `log.py`, `status.py`, ...) plus `core.py` helpers.
- `scripts/perf_check.py` and `perf_budget.json` – performance and structure checks run on every PR.
- `scripts/build_planner.py` – generates the workbooks (needs `openpyxl`; `--recalc` also needs LibreOffice):
  ```
  python scripts/build_planner.py Project_Planner.xlsx --baselined --recalc
  python scripts/build_planner.py Project_Planner_Blank.xlsx --blank --recalc
  ```
  `--recalc` saves the file with calculated values so previews (GitHub, phone file viewers) show numbers; Excel recalculates on open either way.
  CI fails if the committed workbooks don't match the build script, so rebuild and commit both after changing the script.

## Sheets
| Sheet | Purpose |
|---|---|
| Start | The five settings a plan needs (name, manager, start date, working week, hours/day), a 5-step quick start, then the full reference |
| Status | The weekly page, printable on one page (protected without a password, like the Roadmap and the Resource Report; the blue inputs still take typing): the status date and this week's steps (with links), overall RAG and RAG by area (calculated, with PM override and comment), key figures (including the duration-weighted SPI), tasks by status, next milestones, top open RAID items, the first five late tasks, tasks due and starting in the look-ahead and open critical tasks, change control, commentary for the period, sprint load chart, and the snapshot row for History |
| Resource Report | Per-resource work, progress, late / critical tasks; hours and utilisation grids by week and by sprint (over-allocation in red); leave by week; team hours vs capacity chart |
| Agile | Sprint calendar (dates you can override, goal, available person-days, capacity in points after focus %, advisory / overhead days, planned days, committed / done story points, velocity, release burndown), release forecast, a "Sprint to show" picker with that sprint's backlog (tasks, stage, points) and daily burndown, charts |
| Log | One list for risks, assumptions, issues, dependencies, decisions and change requests: Type, owner, dates, status (the dropdown follows the type), linked task; RAID score and rating, decision options and reasons, change schedule and cost impact (columns that don't apply to a row's type are grey); overdue flag |
| History | Status history (paste the Status sheet's snapshot row as values each reporting date) and trend charts: finish, % done vs % planned, SPI, burn-up |
| Settings | Everything else, with defaults: rescheduling, units, custom working days, Live resource grids (No = pause the per-person grids for bulk entry), View from (scrolls the Gantt and Resource Report together), Gantt scale, sprint calendar, RAG thresholds; the holiday table below; the dropdown lists in a collapsed group to the right |
| Tasks | The plan and its Gantt chart on one sheet. Shown: ID, UID, Level, Task Name, Duration, Predecessors (`12, 14SS+2d`, by UID), Resources (`Alex Chen, JL[50%]`), Stage (Not Started … Complete, each with a % on Settings > Stages), actual start / finish, Start and Finish (the current schedule), Sprint, Notes and **Issues** (everything wrong on the row, in words). Collapsed sections (click + above): scheduling (deadline, mode, constraints), detail (status, critical, planned dates, variances, late dates, slack, checks, a % Complete override, WBS, Show), work (resource names, over-allocation), agile (story points, planned sprint), baseline (start, finish, duration, work: set and compare). Money is on the Cost sheet. Each input shows what to type when selected. To the right of the table (link at the top), on the same rows, the **Gantt timeline**: planned dates as a black outline (the baseline once set), the current schedule as fill (done, remaining, critical, red slip past the planned finish), planned days with no work on them now, summaries, milestones, deadlines, status-date line (conditional formatting only, no formulas); day, week, month or quarter scale; outline grouping (1 2 3 buttons) and a Show-to-level filter |
| Roadmap | One-page view for sponsors: each phase with its owner, progress (duration-weighted) and tasks done ("2 of 6 done"), a RAG from its finish slip, its next milestone and a link to its row; the milestones; all on an 18-month timeline (forecast, done, late, planned, missed deadlines) |
| Resources | One row per person: allocation %, start and end dates on the project, focus % (blank = Settings > Focus factor, 80%; 0 for advisory roles), assigned tasks and hours, then their leave by month (`16, 17(0.5)` under `Nov-26`), which lowers weekly and sprint capacity, utilisation and over-allocation (task dates don't move). Below: person-days and story points per sprint (1 person-day = 1 point) |
| Cost | Optional, delete it for audiences that shouldn't see money: day rates (rate card), cost per task (work x day rate, rolled up to summaries), actual cost (work done x rate), budget from the baseline, earned value (BAC, PV, EV, AC, SPI, CPI, EAC, VAC, TCPI) with a cost RAG, task cost vs resource cost (allocated days x rate), cost history and charts. Nothing else reads it |

## Moving to Microsoft Project
**Recommended: Microsoft Project XML.** Save the workbook, then convert it:
- in a browser with [`docs/export.html`](docs/export.html): choose the .xlsx, download the .xml. Nothing to install,
  it works on a phone, and the workbook never leaves your device (the page runs the converter with Pyodide);
- or on the command line: `python scripts/export_mspdi.py Project_Planner.xlsx project.xml`.

Both run the same code (`planner/mspdi.py`). In MS Project: File → Open → the .xml → "Open as a new project". The
file carries the calendar (working week, hours per day, holidays), tasks and outline, links by UID with type and lag,
constraints (all eight types), manual tasks, deadlines, actuals and actual cost (from the Cost sheet, when it is there), the baseline, resources with
rates, assignments with units, and Sprint / Story Points in the custom fields Text1 / Number1.

**From Microsoft Project to the planner.** In MS Project: File → Save As → XML. Then choose that file on the same
page ([`docs/export.html`](docs/export.html)), or run `python scripts/import_mspdi.py project.xml Planner.xlsx`. You
get a new planner workbook with every sheet and report: tasks, outline, links, constraints, deadlines, actuals,
baseline, resources and assignments, calendar, holidays, and resource-calendar days off on the Leave sheet. Anything
the planner can't hold (a fifth predecessor, a lag in % ...) is listed and noted on the task.

**Upgrading to a newer template.** Made your plan with an older version? Choose the workbook on the same page with
"Move it to the latest template" (or run `python scripts/upgrade_planner.py Old.xlsx New.xlsx`). Everything you
typed (tasks, people, assignments, leave, holidays, settings, the Log, status history, sprint
goals, PM RAG and notes) is carried over by column name; the formulas, sheets and reports are the latest.

**MS Project's Import Wizard** needs the `MSP_Tasks` / `MSP_Resources` sheets, which are no longer built by default
(the XML file carries more). Build them with `python scripts/build_planner.py Planner.xlsx --import-wizard-sheets`.

## Limits
- Tables start with 200 task rows, 50 resources, 100 assignments, 26 sprints and 100 holidays, and grow when you add rows. The fixed
  grids (Gantt timeline 150 columns, per-resource grids 25 resources, Resource Report 26 weeks) are sized at
  build time. For a bigger template, rebuild:
  ```
  python scripts/build_planner.py Big_Planner.xlsx --blank --recalc --tasks 2000 --resources 100 --assignments 500 \
      --grid-resources 60 --sprints 52 --gantt-columns 365 --weeks 52
  ```
- Predecessors are one cell per task, by UID: `12, 14SS+2d` (types FS / SS / FF / SF, lag in working days, up to 4).
  Resources are one cell too: names or initials from the Resources sheet, comma-separated, with an optional share of
  their time: `Alex Chen, JL[50%]` (up to 4 people). Each person's work is the task's duration x their units (MS
  Project's fixed units; no effort-driven split), and every per-resource total, grid and over-allocation flag
  includes them. Anything the cells can't read is explained in Issues.
- Insert tasks with right-click → Insert → Table Rows Above (formulas fill in), then give the row `NextFreeUID`.
- **A predecessor must be a detail task listed above its successor.** It keeps the schedule and the critical path consistent. Links that break it (or point to the task itself, an empty row or a missing UID) are ignored and explained in Link Check. Plans imported from MS Project may need reordering.
- The status date is a typed date (Ctrl+; for today), not `TODAY()`, so it doesn't force a full recalculation after every edit. Update it when you report progress.

## Tests
```
python -m unittest discover -s tests -v
```
Builds the workbook, recalculates it with headless LibreOffice (`soffice`) at a fixed status date of 25-Oct-2026,
and checks: no formula errors, WBS numbering, FS / SS / FF links with lags, several predecessors, weekends and holidays skipped,
every constraint type (As Late As Possible, Start / Finish No Earlier / Later Than, Must Start / Finish On), manual tasks, summary roll-ups, forecast from actuals, slips rippling to successors,
rescheduling after the status date, deadlines, frozen baseline and variances, earned value and the RAG rules (worked out independently in Python), PM RAG overrides, the report lists, velocity and burndown, the planned-sprint override, weekly resource hours (recomputed in Python) and over-allocation, RAID scoring and links, change control totals, status snapshots, invalid-link handling, inserting a task
mid-list (links follow UIDs), the blank template, sprint tagging, cost, the MS Project export columns, and that the
committed workbooks match the build script, tables keep their calculated columns, and no formula contains an A1
reference. GitHub Actions runs the same suite on every push, then `scripts/perf_check.py` on the sample and on a
1,000-task stress workbook (formula count, volatile functions, recalculation time against `perf_budget.json`).
Recalculation tests are skipped if LibreOffice is not installed.
