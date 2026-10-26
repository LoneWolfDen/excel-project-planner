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
   Excel recalculate them and everything downstream after every edit. The only one allowed is `OFFSET`, in the seven
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
| Status Report | `OverallStatus`, `StatusRAG`, `EarnedValue`, `StatusNotes` | RAG by area, earned value (`evm_*` names) |
| Schedule Report | `PhaseSummary`, `MilestoneReport`, `LateTasks`, `StartingSoon`, `DueSoon`, `CriticalTasks` | Ranked lists |
| Resource Report | `ResourceSummary` + 2 week grids | Weekly hours (`rw_Load`) and utilisation (`rw_Util`) per resource |
| Agile Report | `AgileKPIs`, `SprintBacklog`, `SprintBurndown` | Velocity forecast (`agl_*` names), current sprint |
| RAID / Decisions / Changes | `RAIDLog`, `DecisionLog`, `ChangeLog` | Score, rating, overdue, linked task (`raid_*`, `dec_*`, `chg_*` names) |
| History | `SnapshotNow`, `StatusHistory` | This status date's numbers; pasted history for trend charts |
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

`OFFSET` is therefore used in exactly seven columns: the four predecessor drivers, the parent-WBS look-up and the critical path's two successor look-ups (`_SuccessorFinishLimit`, `Free Slack`). Volatile
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

## Critical path (backward pass)

Each successor row works out, for each of its two links, the latest its predecessor may finish:
FS `WORKDAY(LateStart, -(1 + lag))` (no +1 into a milestone), FF `WORKDAY(LateFinish, -lag)`, and SS / SF the latest
start converted to a finish with the predecessor's duration (an input, so no circularity). These mirror the forward
pass exactly.

The predecessor takes the earliest of those limits, its deadline and `ProjectFinish` (the latest forecast finish).
Successors sit below their predecessors, so it looks only at the rows from the one below down to
`_LastSuccessorRow`. That helper reads only input columns, so it isn't volatile, and it keeps each look-up short.
Limits are stored as keys (`31-Dec-9999 - date`, `100000 - slack`) so that one `MAXIFS` per link answers both
"is there a successor" and "what is its tightest limit".

Total Slack = working days from forecast finish to late finish; Free Slack = the smallest slack on any outgoing
link (capped at total slack); Critical = an unfinished detail task with total slack <= `CriticalSlackDays`. A missed
deadline gives negative slack to the task and the chain feeding it, as in MS Project.

| LibreOffice, one edit then recalculate | Before the critical path | With it |
|---|---|---|
| Sample (200 rows) | 0.04 s | 0.05 s |
| 1,000 tasks | 0.42 s | ~0.5 s |

## Reports

Reports are formulas, not Power Query, so they are current after every edit, work in Excel for the web and on a
phone, and are tested in CI. They follow three rules:

- **Tasks decides, the report looks up.** A list such as "late tasks" needs a filter and a sort. Tasks does both in
  one helper column per list (`_LateKey`, `_StartSoonKey`, `_CriticalKey`, ...): blank when the task is not on the
  list, otherwise a unique sort key (the row position breaks ties). Row *k* of the report finds the *k*-th key with
  `SMALL` / `LARGE` and `MATCH`, once, in its hidden `_Task` column; the visible columns only `INDEX` by it
  (`core.ranked_table`). No `FILTER` / `SORT` (Microsoft 365 only) and no `SUMPRODUCT` per row.
- **Totals sum leaf helpers.** Earned value is worked out per detail task (`_LeafBudget`, `_LeafPV`, `_LeafEV`,
  `_LeafAC`, zero on summaries), so a project total is one `SUM` and a phase total is the summary row's roll-up.
- **One rule, one place.** RAG thresholds are settings; `planner/rag.py` turns a measure into a severity (0 / 1 / 2)
  and the worst severity into a word, for both the status rows and the phase table. A PM RAG typed next to a
  calculated RAG overrides it.

Earned value at the status date: Budget = Baseline Cost (or Cost before a baseline); planned % = planned working days
before the status date / planned working days; PV = Budget x planned %; EV = Budget x Progress; AC = Actual Cost, or
Cost x Progress when blank; EAC = AC + (BAC - EV) / CPI.

**History without macros.** A formula can't remember last week's value, so status history works like the baseline:
`SnapshotNow` is one row of formulas (names only) and the PM pastes it as values into `StatusHistory`, an input
table the trend charts read. A macro or script may automate the copy, but the workbook never depends on it.

Weekly resource hours use the same SUMIFS spreading as the sprint load (`sprints.work_between`), anchored on
`rw_LoadOrigin`; story points are done on a task's Actual Finish (`_DoneDate`), and per-sprint points are columns
of `SprintTable`, so the release burndown is one formula per column.

| LibreOffice, one edit then recalculate | Before the reports | Status + Schedule | + Resource + Agile |
|---|---|---|---|
| Sample (200 rows) | 0.05 s | 0.06 s | 0.08-0.10 s |
| 1,000 tasks | 0.54 s | 0.57-0.61 s | 0.68-0.79 s |

The weekly resource grids first cost 0.25 s per edit at 1,000 tasks (six SUMIFS per cell). A hidden grid of
cumulative hours at each week boundary (three SUMIFS, none outside the project's dates) brought that to about
0.05 s; each week is then the difference of two neighbours.

## Excel version

**Excel 2019 or Microsoft 365** (Windows or Mac). The only post-2010 functions are `MINIFS` and `MAXIFS`, stored with
the `_xlfn.` prefix Excel uses. Microsoft 365-only functions (`XLOOKUP`, `FILTER`, `LET`, `LAMBDA`, ...) are not used: a
unit test enforces this, and LibreOffice, which runs the tests, couldn't check them.

## Guardrails

Run on every PR by `.github/workflows/ci.yml`:

| Check | Where | Fails when |
|---|---|---|
| Unit tests: schedule, baseline, links, inserting rows, blank template | `tests/test_planner.py` | any result differs from the hand-worked dates |
| No A1 references in formulas | `scripts/perf_check.py`, also a unit test | any formula contains one |
| Excel limits: balanced parentheses, length <= 8,192, nesting <= 64 | `perf_check.py`, also a unit test | any formula breaks one (LibreOffice silently repairs some; Excel would not) |
| Real row insertion and deletion keep every formula working | `RealRowEditTest` (LibreOffice via UNO) | a task can't see an inserted row, or a deleted row leaves `#REF!` |
| Volatile functions only on the allow-list (OFFSET in 7 Tasks columns) | `perf_check.py`, `perf_budget.json`, also a unit test | any other volatile use appears |
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
