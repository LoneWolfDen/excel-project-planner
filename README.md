# Excel Project Planner (MS Project-compatible)

An Excel workbook that works like Microsoft Project: tasks and sub-tasks, dependencies,
planned / actual / forecast dates, a Gantt chart view, resources and sprint capacity.
Formulas use named ranges (`ProjectStart`, `tsk_PlanStart`, `res_Rate`, ...) instead of raw cell references.

## Files
- `Project_Planner.xlsx` – the template, loaded with a 30-task sample project.
- `docs/` – [MS Project gap analysis](docs/MSP_GAP_ANALYSIS.md), [backlog](docs/BACKLOG.md), [testing on iPhone](docs/TESTING_ON_IPHONE.md).
- `scripts/build_planner.py` – regenerates the workbook: `python scripts/build_planner.py Project_Planner.xlsx` (needs `openpyxl`).

## Sheets
| Sheet | Purpose |
|---|---|
| Instructions | How to use it, colour key, MS Project import steps, named ranges |
| Dashboard | Planned vs forecast finish, slip, progress, hours, cost, next milestones, charts |
| Settings | Project start, status date, working week, hours/day, Gantt scale, sprint calendar |
| Tasks | Entry table: Level, Name, Mode, Duration, 2 predecessors (FS/SS/FF/SF + lag), actuals, % complete, resource |
| Gantt | Planned line, done / remaining / late bars, summaries, milestones, status-date line |
| Resources | Resource sheet with max units and rates; assigned hours and cost |
| Sprints | Sprint calendar, capacity vs planned hours, per-person load with over-allocation flags |
| Holidays | Non-working dates used by every date formula |
| MSP_Tasks / MSP_Resources | Columns named with MS Project field names for the Import Wizard |

## Importing into Microsoft Project
1. File → Options → Trust Center → Legacy Formats → allow loading legacy/non-default formats.
2. File → Open → this workbook → Import Wizard → new map → tick Tasks and Resources, using the `MSP_Tasks` and `MSP_Resources` sheets.
3. For auto-scheduled tasks leave Start/Finish unmapped (otherwise MS Project adds date constraints); map Baseline, Actual and % Complete fields.

## Limits
- 200 task rows, 50 resources, 26 sprints, 100 holidays (extend rows and named ranges in Name Manager).
- Task IDs renumber when rows are inserted; predecessor IDs you typed do not update.
- Link detail tasks, not summary tasks, to avoid circular references.

## Tests
```
python -m unittest discover -s tests -v
```
Builds the workbook, recalculates it with headless LibreOffice (`soffice`) at a fixed status date of 25-Oct-2026,
and checks: no formula errors, WBS numbering, FS / SS / FF links with lags, two predecessors, weekends and holidays skipped,
Start No Earlier Than, manual tasks, summary roll-ups, forecast from actuals, slips rippling to successors,
rescheduling after the status date, sprint tagging, cost, and the MS Project export columns.
Recalculation tests are skipped if LibreOffice is not installed.
