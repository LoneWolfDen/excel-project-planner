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
- `docs/` – [architecture and rules](docs/ARCHITECTURE.md), [MS Project gap analysis](docs/MSP_GAP_ANALYSIS.md), [backlog](docs/BACKLOG.md), [testing on iPhone](docs/TESTING_ON_IPHONE.md).
- `planner/` – the generator: one module per sheet (`tasks.py`, `gantt.py`, `sprints.py`, ...) plus `core.py` helpers.
- `scripts/perf_check.py` and `perf_budget.json` – performance and structure checks run on every PR.
- `scripts/build_planner.py` – generates the workbooks (needs `openpyxl`; `--recalc` also needs LibreOffice):
  ```
  python scripts/build_planner.py Project_Planner.xlsx --recalc
  python scripts/build_planner.py Project_Planner_Blank.xlsx --blank --recalc
  ```
  `--recalc` saves the file with calculated values so previews (GitHub, phone file viewers) show numbers; Excel recalculates on open either way.
  CI fails if the committed workbooks don't match the build script, so rebuild and commit both after changing the script.

## Sheets
| Sheet | Purpose |
|---|---|
| Instructions | How to use it, colour key, MS Project import steps, named ranges |
| Dashboard | Planned vs forecast finish, slip, progress, hours, cost, next milestones, charts |
| Settings | Project start, status date, working week, hours/day, Gantt scale, sprint calendar |
| Tasks | Entry table: UID, Level, Name, Mode, Duration, 2 predecessors by UID (FS/SS/FF/SF + lag), deadline, actuals, % complete, resource; frozen baseline with start / finish / duration / work / cost variance; critical path (late start/finish, total and free slack, Critical); Link Check and Deadline Check columns |
| Gantt | Planned outline, done / remaining / critical / late bars, summaries, milestones, deadlines, status-date line (conditional formatting only, no formulas) |
| Resources | Resource sheet with max units and rates; assigned hours and cost |
| Sprints | Sprint calendar, capacity vs planned hours, per-person load with over-allocation flags |
| Holidays | Non-working dates used by every date formula |
| MSP_Tasks / MSP_Resources | Columns named with MS Project field names for the Import Wizard |

## Importing into Microsoft Project
1. File → Options → Trust Center → Legacy Formats → allow loading legacy/non-default formats.
2. File → Open → this workbook → Import Wizard → new map → tick Tasks and Resources, using the `MSP_Tasks` and `MSP_Resources` sheets.
3. For auto-scheduled tasks leave Start/Finish unmapped (otherwise MS Project adds date constraints); map Baseline, Actual and % Complete fields.
4. Map `Predecessors` (by ID) for a new project. `Unique ID` and `Unique ID Predecessors` are there for merging into an existing project.

## Limits
- Tables start with 200 task rows, 50 resources, 26 sprints and 100 holidays, and grow when you add rows. The Gantt and MSP_Tasks sheets show the first 200 tasks (rebuild with `--tasks N` for more).
- Insert tasks with right-click → Insert → Table Rows Above (formulas fill in), then give the row `NextFreeUID`.
- **A predecessor must be a detail task listed above its successor.** It keeps the schedule and the critical path consistent. Links that break it (or point to the task itself, an empty row or a missing UID) are ignored and explained in Link Check. Plans imported from MS Project may need reordering.
- The status date is a typed date (Ctrl+; for today), not `TODAY()`, so it doesn't force a full recalculation after every edit. Update it when you report progress.

## Tests
```
python -m unittest discover -s tests -v
```
Builds the workbook, recalculates it with headless LibreOffice (`soffice`) at a fixed status date of 25-Oct-2026,
and checks: no formula errors, WBS numbering, FS / SS / FF links with lags, two predecessors, weekends and holidays skipped,
Start No Earlier Than, manual tasks, summary roll-ups, forecast from actuals, slips rippling to successors,
rescheduling after the status date, deadlines, frozen baseline and variances, invalid-link handling, inserting a task
mid-list (links follow UIDs), the blank template, sprint tagging, cost, the MS Project export columns, and that the
committed workbooks match the build script, tables keep their calculated columns, and no formula contains an A1
reference. GitHub Actions runs the same suite on every push, then `scripts/perf_check.py` on the sample and on a
1,000-task stress workbook (formula count, volatile functions, recalculation time against `perf_budget.json`).
Recalculation tests are skipped if LibreOffice is not installed.
