# Excel Project Planner (MS Project-compatible)

An Excel workbook that works like Microsoft Project: tasks and sub-tasks, dependencies,
planned / actual / forecast dates, a Gantt chart view, resources and sprint capacity.
Formulas use named ranges (`ProjectStart`, `tsk_PlanStart`, `res_Rate`, ...) instead of raw cell references.

## Files
- `Project_Planner.xlsx` – the template, loaded with a 30-task sample project.
- `Project_Planner_Blank.xlsx` – the same template with no sample data, for a new project.
- `docs/` – [MS Project gap analysis](docs/MSP_GAP_ANALYSIS.md), [backlog](docs/BACKLOG.md), [testing on iPhone](docs/TESTING_ON_IPHONE.md).
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
| Tasks | Entry table: UID, Level, Name, Mode, Duration, 2 predecessors by UID (FS/SS/FF/SF + lag), deadline, actuals, % complete, resource; frozen baseline with start / finish / duration / work / cost variance; Link Check and Deadline Check columns |
| Gantt | Planned line, done / remaining / late bars, summaries, milestones, deadline markers, status-date line |
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
- 200 task rows, 50 resources, 26 sprints, 100 holidays (extend rows and named ranges in Name Manager).
- Insert tasks with *Insert Copied Cells* of an empty task row (a plain inserted row has no formulas), then give it `NextFreeUID`.
- Links to the task itself, its own summary, an empty row or a missing UID are ignored and explained in Link Check; a loop through several tasks (A → B → A) is still a circular reference.

## Tests
```
python -m unittest discover -s tests -v
```
Builds the workbook, recalculates it with headless LibreOffice (`soffice`) at a fixed status date of 25-Oct-2026,
and checks: no formula errors, WBS numbering, FS / SS / FF links with lags, two predecessors, weekends and holidays skipped,
Start No Earlier Than, manual tasks, summary roll-ups, forecast from actuals, slips rippling to successors,
rescheduling after the status date, deadlines, frozen baseline and variances, invalid-link handling, inserting a task
mid-list (links follow UIDs), the blank template, sprint tagging, cost, the MS Project export columns, and that the
committed workbooks match the build script. GitHub Actions runs the same suite on every push.
Recalculation tests are skipped if LibreOffice is not installed.
