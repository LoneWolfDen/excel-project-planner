# Architecture

How the workbook is put together, and the rules every change follows. These rules are checked automatically
(see [Guardrails](#guardrails)); a PR that breaks one fails CI.

## Principles

1. **No A1 addresses in formulas.** Every formula refers to cells through
   - an **Excel Table** column: `Tasks[[#This Row],[Duration]]`, which Excel shows as `[@Duration]`, or `Tasks[Duration]`;
   - a **defined name**: `ProjectStart`, `HoursPerDay`, `tsk_Duration` (= `Tasks[Duration]`);
   - a **position relative to a named anchor**: `ROW()-ROW(Tasks[[#Headers],[ID]])`, `COLUMN()-COLUMN(gantt_Origin)`.

   Inserting, deleting or moving rows and columns therefore never breaks a formula, and nothing needs re-pointing.
   The only A1 ranges are where Excel allows nothing else: conditional formatting, data validation and chart
   series. Excel shifts those itself when rows or columns are inserted.
2. **One formula per column.** Inside a table every row of a column holds the same formula text: a calculated
   column. Excel fills it into inserted rows automatically. Grids outside tables (Sprints load and utilisation,
   Gantt header dates) also use one formula per grid, located by `ROW()`/`COLUMN()` offsets.
3. **Inputs and formulas never mix.** Blue text = input, grey fill = formula, light-grey header = hidden helper.
   Summaries roll up rather than reading inputs.
4. **Readable formulas.** Columns have short, plain names (`Forecast Finish`, `Pred1 UID`); helper columns start
   with `_` (`_CalcStart`, `_Pred1Row`). Each helper does one job so the visible formulas stay short. Header comments
   explain units and meaning.
5. **Cheap to recalculate.** Volatile functions are allowed only where listed in `perf_budget.json`. No whole-column
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

## How scheduling avoids circular references

A task's start depends on another row's finish, and a summary's dates depend on the rows below it. Written as
`INDEX(Tasks[_CalcFinish], ...)` either would look circular to Excel, because Excel judges circularity from the
ranges a formula names and that range contains the formula's own row. So these look-ups use
`OFFSET(Tasks[[#Headers],[_CalcFinish]], row, 0)`. Its only fixed precedent is the header cell, and Excel resolves
the target at calculation time.

The price is that `OFFSET` is **volatile**: after every edit Excel recalculates those cells and everything that
depends on them, which is most of the workbook. The measured cost:

| Workbook | Formulas | Full recalculation | One edit |
|---|---|---|---|
| Sample, 200 task rows | 21,758 | 0.085 s | 0.061 s |
| Stress, 1,000 real tasks | 97,758 | ~0.6 s | ~0.45 s |

LibreOffice 24.2, single machine; Excel is usually faster. Before this design (v1.1) the sample had 50,932 formulas,
and a 1,000-task stress edit took ~0.86 s.

A non-volatile design is possible if every predecessor sits **above** its successor and links go to detail tasks
only. Look-ups can then use plain ranges ending at the row above, and an edit recalculates only what depends on it.
It is a behaviour change, so it is a separate decision (see the backlog).

## Guardrails

Run on every PR by `.github/workflows/ci.yml`:

| Check | Where | Fails when |
|---|---|---|
| Unit tests: schedule, baseline, links, inserting rows, blank template | `tests/test_planner.py` | any result differs from the hand-worked dates |
| No A1 references in formulas | `scripts/perf_check.py`, also a unit test | any formula contains one |
| Volatile functions only on the allow-list | `perf_check.py`, `perf_budget.json` | a new volatile use appears |
| Formula count | `perf_check.py` | above budget |
| Full and single-edit recalculation time, sample and 1,000-task stress workbook | `perf_check.py --timing` (LibreOffice) | above budget |
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
