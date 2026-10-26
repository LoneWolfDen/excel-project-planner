# Architecture

How the workbook is put together, and the rules every change follows. These rules are checked automatically
(see [Guardrails](#guardrails)); a PR that breaks one fails CI.

## Principles

1. **No A1 addresses in formulas.** Every formula refers to cells through
   - an **Excel Table** column: `Tasks[[#This Row],[Duration]]`, which Excel shows as `[@Duration]`, or `Tasks[Duration]`;
   - a **defined name**: `ProjectStart`, `HoursPerDay`, `tsk_Duration` (= `Tasks[Duration]`);
   - a **position relative to a named anchor**: `ROW()-ROW(Tasks[[#Headers],[ID]])`, `COLUMN()-COLUMN(gantt_Origin)`.

   Inserting, deleting or moving rows and columns therefore never breaks a formula, and nothing needs re-pointing.
   A1 ranges appear only where Excel allows nothing else:
   - conditional formatting, data validation and chart series;
   - **rows-above ranges** inside the Tasks table: `$X$5:X9` in row 10, from the column's header cell to the row
     above. Excel has no structured reference for "the rows above this one", and the schedule needs exactly that
     (see below). Excel keeps these correct when rows are inserted or deleted; `perf_check.py` allows this one
     shape and nothing else.
2. **One formula per column.** Inside a table every row of a column holds the same formula text: a calculated
   column. Excel fills it into inserted rows automatically. Grids outside tables (Sprints load and utilisation,
   Gantt header dates) also use one formula per grid, located by `ROW()`/`COLUMN()` offsets.
3. **Inputs and formulas never mix.** Blue text = input, grey fill = formula, light-grey header = hidden helper.
   Summaries roll up rather than reading inputs.
4. **Readable formulas.** Columns have short, plain names (`Forecast Finish`, `Pred1 UID`); helper columns start
   with `_` (`_CalcStart`, `_Pred1Row`). Each helper does one job so the visible formulas stay short. Header comments
   explain units and meaning.
5. **Cheap to recalculate.** No volatile functions (`OFFSET`, `INDIRECT`, `TODAY`, `NOW`, `RAND`, `CELL`, `INFO`):
   Excel recalculates a volatile cell and everything downstream after every edit. Dates the user would expect to
   be "today" (the status date) are typed instead. No whole-column
   (`A:A`) references, and no formulas in cells that only exist to be coloured: the Gantt timeline is conditional
   formatting only. A lookup used by several columns is done once, in a helper column.
6. **Generated, not hand-edited.** The workbooks are built by `planner/` (Python + openpyxl). Change the code, then
   rebuild; CI fails if a committed `.xlsx` doesn't match the code.

## Layout

| Sheet | Table(s) | What is calculated there |
|---|---|---|
| Settings | (named cells) | Project start, status date, calendar mask (`WorkWeek`), Gantt and sprint settings |
| Tasks | `Tasks` | All scheduling: WBS, summaries, links, planned / forecast dates, variances, progress, work, cost |
| Gantt | `GanttView` + timeline | Task panel reads Tasks by position; timeline = conditional formatting only |
| Resources | `ResourceTable` | Assigned tasks, work, cost per resource |
| Sprints | `SprintTable` + 2 grids | Sprint calendar, capacity, load and utilisation per resource |
| Holidays | `HolidayTable` | Non-working dates (`Holidays` = `HolidayTable[Date]`) |
| MSP_Tasks / MSP_Resources | `MSPTasks`, `MSPResources` | MS Project field names, for the Import Wizard |
| Dashboard | `KPIs`, `StatusSummary`, `NextMilestones` | Project KPIs |
| Lists | `List*` | Dropdown values (`L_DepType`, `L_Mode`, ...) |

Code: `planner/core.py` (styles, `add_table`, structured-reference helpers), one module per sheet, `planner/build.py`
(command line, LibreOffice recalculation). `scripts/build_planner.py` is the entry point.

## How scheduling avoids circular references without volatile functions

A task's start depends on another row's finish, and a summary's dates depend on the rows below it. Excel judges
circularity from the *ranges* a formula names, not the cell a look-up picks, so `INDEX(Tasks[Finish], n)` inside
the Finish chain counts as circular even when n points elsewhere. Two rules remove the problem:

1. **A predecessor sits above its successor and is a detail task.** Predecessor look-ups read only the rows above
   (`INDEX($X$5:X9, pos)` in row 10), which can never include the row itself or anything below it. Link Check
   explains a link that breaks the rule ("UID 21 is below this task: move it above", "UID 15 is a summary: link to
   one of its tasks"), and the link is ignored rather than causing an error. MS Project allows both, so a plan
   imported from it may need reordering.
2. **Detail values and roll-ups live in different columns.** `_LeafCalcStart`, `_LeafFcstFinish`, `_LeafWork`, ... hold
   detail-task values; summary rows roll them up with `MIN(INDEX(col, first):INDEX(col, last))`. The leaf columns
   never read the roll-up columns.

The result has no volatile cells, so Excel recalculates only what an edit can affect. Excel tracks dependencies by
range, so changing a task's dates recalculates the tasks below it, the summaries and the reports. Edits that don't
touch dates (names, notes, resources, settings that only drive views) recalculate almost nothing.

| LibreOffice, one edit then recalculate | OFFSET design (PR #69) | This design |
|---|---|---|
| Sample (200 rows): duration near the top | 0.061 s | 0.046 s |
| 1,000 tasks: duration near the top (worst case) | 0.46 s | 0.26 s |
| 1,000 tasks: last task's duration | 0.46 s | 0.23 s |
| 1,000 tasks: a note | ~0.5 s | 0.003 s |
| Volatile cells (1,000 tasks) | 12,002 | 0 |

Sprint and resource load use the cumulative form `W(D) = SUMIFS(work, finish < D) + D x SUMIFS(work per day,
running on D) - SUMIFS(work before start, running on D)` in working days since the project start, so a load cell
is six `SUMIFS` rather than a `SUMPRODUCT` over every task, and weekends and holidays carry no work.

## Guardrails

Run on every PR by `.github/workflows/ci.yml`:

| Check | Where | Fails when |
|---|---|---|
| Unit tests: schedule, baseline, links, inserting rows, blank template | `tests/test_planner.py` | any result differs from the hand-worked dates |
| No A1 references in formulas (rows-above ranges in tables excepted) | `scripts/perf_check.py`, also a unit test | any other A1 reference appears |
| No volatile functions | `perf_check.py`, `perf_budget.json`, also a unit test | any volatile function appears |
| Formula count | `perf_check.py` | above budget |
| Recalculation time: full, a duration edit near the top, the last task, a note; sample and 1,000-task stress workbook | `perf_check.py --timing` (LibreOffice) | above budget |
| Committed workbooks match the build | unit test | the `.xlsx` wasn't rebuilt after a code change |
| Tables keep their calculated-column formulas | unit test | Excel wouldn't auto-fill inserted rows |

The results appear in the CI run summary. A budget may be raised only with a reason in the PR description.

Run the same checks locally:

```
python -m unittest discover -s tests -v
python scripts/build_planner.py /tmp/stress.xlsx --tasks 1000 --stress 33
python scripts/perf_check.py Project_Planner.xlsx --profile sample --timing
python scripts/perf_check.py /tmp/stress.xlsx --profile stress --timing
```

`--timing` needs LibreOffice and `python3-uno`.
