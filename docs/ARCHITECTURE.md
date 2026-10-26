# Architecture

How the workbook is put together, and the rules every change follows. These rules are checked automatically
(see [Guardrails](#guardrails)); a PR that breaks one fails CI.

## Principles

1. **No A1 addresses in formulas.** Every formula refers to cells through
   - an **Excel Table** column: `Tasks[[#This Row],[Duration]]`, which Excel shows as `[@Duration]`, or `Tasks[Duration]`;
   - a **defined name**: `ProjectStart`, `HoursPerDay`, `tsk_Duration` (= `Tasks[Duration]`);
   - a **position relative to a named anchor**: `ROW()-ROW(Tasks[[#Headers],[ID]])`, `COLUMN()-COLUMN(gantt_Origin)`.

   Inserting, deleting or moving rows and columns therefore never breaks a formula, and nothing needs re-pointing.
   A1 ranges appear only where Excel allows nothing else: conditional formatting, data validation and chart
   series. Excel shifts those itself when rows or columns are inserted.
2. **One formula per column.** Inside a table every row of a column holds the same formula text: a calculated
   column. Excel fills it into inserted rows automatically. Grids outside tables (Sprints load and utilisation,
   Gantt header dates) also use one formula per grid, located by `ROW()`/`COLUMN()` offsets.
3. **Inputs and formulas never mix.** Blue text = input, grey fill = formula, light-grey header = hidden helper.
   Summaries roll up rather than reading inputs.
4. **Readable formulas.** Columns have short, plain names (`Forecast Finish`, `Pred1 UID`); helper columns start
   with `_` (`_CalcStart`, `_Pred1Row`). Each helper does one job so the visible formulas stay short. Header comments
   explain units and meaning.
5. **Cheap to recalculate.** Volatile functions (`OFFSET`, `INDIRECT`, `TODAY`, `NOW`, `RAND`, `CELL`, `INFO`) make
   Excel recalculate them and everything downstream after every edit. The only one allowed is `OFFSET`, in the five
   Tasks columns that must read another row's calculated value (below). Dates you'd expect to be "today" (the
   status date) are typed instead. Report grids use `SUMIFS`, not `SUMPRODUCT` over every task. No whole-column
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

## Reading another row's calculated value

A task's start depends on its predecessor's finish, and the WBS number on the parent's WBS: one row reads a value
*calculated* in another row of the same table. Excel offers three ways to do that, and only one is safe:

| Approach | Circular? | Survives inserting a row directly above? | Volatile? |
|---|---|---|---|
| `INDEX(Tasks[_LeafCalcFinish], pos)` | **Yes**: Excel judges circularity by the range named, which includes the row itself | yes | no |
| `INDEX($X$5:X9, pos)` in row 10 (rows above) | no | **No**: a range ending at the row above does not grow when a row is inserted directly above, so the next task can't see the new one | no |
| `OFFSET(Tasks[[#Headers],[_LeafCalcFinish]], pos, 0)` | no | yes | **yes** |

The rows-above design was built and measured first (it is fast). `RealRowEditTest`, which inserts a real row with
LibreOffice, showed the inserted task invisible to the task below it, so it was rejected: robustness first.

`OFFSET` is therefore used in exactly five columns: the four predecessor drivers and the parent-WBS look-up. Volatile
cells recalculate, with everything downstream, after every edit. Everything else is non-volatile:

- **Roll-ups.** Detail-task values live in `_Leaf...` columns (`_LeafCalcStart`, `_LeafFcstFinish`, `_LeafWork`, ...);
  summary rows roll them up with `MIN(INDEX(col, first):INDEX(col, last))`, and the leaf columns never read the
  roll-up columns.
- **Status date** is a typed date, not `TODAY()`.
- **Sprint and resource load** use the cumulative form `W(D) = SUMIFS(work, finish < D) + D x SUMIFS(work per day,
  running on D) - SUMIFS(work before start, running on D)` in working days since the project start: six `SUMIFS`
  per cell instead of a `SUMPRODUCT` over every task, and weekends and holidays carry no work.

**Ordering rule.** A predecessor must be a detail task listed above its successor. The critical path's backward
pass reads successors from the rows below, so the rule keeps the forward and backward passes consistent. Link
Check explains a link that breaks it, and the link is ignored rather than causing an error.

| LibreOffice, one edit then recalculate | v1.1 | Now |
|---|---|---|
| Sample (200 rows) | 0.09 s | 0.04 s |
| 1,000 tasks | 0.86 s | 0.42 s |
| Volatile cells (sample) | 2,402 | 1,000 |

Excel on Windows is usually several times faster than LibreOffice and calculates on all cores.

## Guardrails

Run on every PR by `.github/workflows/ci.yml`:

| Check | Where | Fails when |
|---|---|---|
| Unit tests: schedule, baseline, links, inserting rows, blank template | `tests/test_planner.py` | any result differs from the hand-worked dates |
| No A1 references in formulas | `scripts/perf_check.py`, also a unit test | any formula contains one |
| Excel limits: balanced parentheses, length <= 8,192, nesting <= 64 | `perf_check.py`, also a unit test | any formula breaks one (LibreOffice silently repairs some; Excel would not) |
| Real row insertion and deletion keep every formula working | `RealRowEditTest` (LibreOffice via UNO) | a task can't see an inserted row, or a deleted row leaves `#REF!` |
| Volatile functions only on the allow-list (OFFSET in 5 Tasks columns) | `perf_check.py`, `perf_budget.json`, also a unit test | any other volatile use appears |
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
