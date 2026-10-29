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
   column. Excel fills it into inserted rows automatically. Grids outside tables (Resource Report load and utilisation,
   Gantt header dates) also use one formula per grid, located by `ROW()`/`COLUMN()` offsets.
3. **Inputs and formulas never mix.** Blue text = input, grey fill = formula, light-grey header = hidden helper.
   Summaries roll up rather than reading inputs.
4. **Readable formulas.** Columns have short, plain names (`Finish`, `Predecessors`); helper columns start
   with `_` (`_CalcStart`, `_Pred1Pos`). Each helper does one job so the visible formulas stay short. Header comments
   explain units and meaning.
5. **Cheap to recalculate.** Volatile functions (`OFFSET`, `INDIRECT`, `TODAY`, `NOW`, `RAND`, `CELL`, `INFO`) make
   Excel recalculate them and everything downstream after every edit. The only one allowed is `OFFSET`, in the eleven
   Tasks columns that must read another row's calculated value (below). Dates you'd expect to be "today" (the
   status date) are typed instead. Report grids use `SUMIFS`, not `SUMPRODUCT` over every task. No whole-column
   (`A:A`) references, and no formulas in cells that only exist to be coloured: the Gantt timeline is conditional
   formatting only. A lookup used by several columns is done once, in a helper column.
6. **Generated, not hand-edited.** The workbooks are built by `planner/` (Python + openpyxl). Change the code, then
   rebuild; CI fails if a committed `.xlsx` doesn't match the code.

## Layout

| Sheet | Table(s) | What is calculated there |
|---|---|---|
| Start | (named cells) | The five essential settings (`ESSENTIAL` in `planner/settings.py`), quick start, reference |
| Settings | (named cells) | The other settings, calendar mask (`WorkWeek`, from the Working week preset or the seven day rows), Gantt and sprint settings |
| Tasks | `Tasks` + timeline | All scheduling: WBS, summaries, links, planned / forecast dates, variances, progress, work; the Gantt timeline to the right of the table (conditional formatting only). Columns are in `SECTIONS`: the essentials, then one collapsed Excel column group per optional section, each with one visible key column (Excel merges groups that touch); the helpers are a level-2 group |
| Resources | `ResourceTable`, `cap_Days`, `cap_Points` | Allocation, dates and focus per person, assigned tasks and work, then one leave column per month (`leave_Grid` / `leave_Months` name only the month columns, so the leave look-ups never depend on the calculated columns) |
| Settings (below the settings) | `HolidayTable` | Non-working dates (`Holidays` = `HolidayTable[Date]`) |
| MSP_Tasks / MSP_Resources | `MSPTasks`, `MSPResources` | MS Project field names, for the Import Wizard; only with `--import-wizard-sheets` |
| Status | `OverallStatus`, `StatusRAG`, `KPIs`, `StatusSummary`, `NextMilestones`, `TopRAID`, `LateTasks`, `DueSoon`, `StartingSoon`, `CriticalTasks`, `ChangeControl`, `StatusNotes`, `SnapshotNow` | The weekly page: status date (`StatusDateInput`), RAG by area, key figures (`sch_Progress`, `sch_Planned`, `sch_SPI`), the snapshot row |
| Cost | `RateCard`, `CostSummary`, `CostSnapshot`, `CostHistory`, `TaskCost` | Optional: day rates (`cost_Rate`), cost per task, earned value (`evm_*`), cost RAG; no other sheet reads it |
| Resource Report | `ResourceSummary` + week and sprint grids | Hours (`rw_Load`, `load_Grid`) and utilisation (`rw_Util`) per resource by week and by sprint, leave by week (`rw_Leave`) and by sprint (`spr_LeaveGrid`, hidden) |
| Agile | `SprintTable`, `AgileKPIs`, `SprintBacklog`, `SprintBurndown` | Sprint calendar, capacity, points and velocity (`spr_*`), release forecast (`agl_*`), current sprint |
| Log | `Log` | RAID items, decisions and change requests: score, rating, overdue, waiting, linked task (`log_*` names; `log_Open` = 1 while open) |
| History | `StatusHistory` | Pasted snapshots for the trend charts |
| Settings (columns G onward, collapsed) | `List*` | Dropdown values (`L_DepType`, `L_Mode`, ...) |

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

`OFFSET` is therefore used in exactly eleven columns: the eight predecessor drivers (plan and forecast for each of the four links), the parent-WBS look-up and the critical path's two successor look-ups (`_SuccessorFinishLimit`, `Free Slack`). Volatile
cells recalculate, with everything downstream, after every edit. Everything else is non-volatile:

- **Roll-ups.** Detail-task values live in `_Leaf...` columns (`_LeafCalcStart`, `_LeafFcstFinish`, `_LeafWork`, ...);
  summary rows roll them up with `MIN(INDEX(col, first):INDEX(col, last))`, and the leaf columns never read the
  roll-up columns.
- **Status date** is a typed date, not `TODAY()`.
- **Sprint and resource load** use the cumulative form `W(D) = SUMIFS(work, finish < D) + D x SUMIFS(work per day,
  running on D) - SUMIFS(work before start, running on D)` in working days since the project start: six `SUMIFS`
  per cell instead of a `SUMPRODUCT` over every task, and weekends and holidays carry no work.

**Ordering rule** ([SCHEDULING_SPEC.md](SCHEDULING_SPEC.md) section 2). A predecessor must be a detail task listed above its successor. The critical path's backward
pass reads successors from the rows below, so the rule keeps the forward and backward passes consistent. Link
Check explains a link that breaks it, and the link is ignored rather than causing an error.

| LibreOffice, one edit then recalculate | v1.1 | Now |
|---|---|---|
| Sample (200 rows) | 0.09 s | 0.04 s |
| 1,000 tasks | 0.86 s | 0.42 s |
| Volatile cells (sample) | 2,402 | 1,000 |

Excel on Windows is usually several times faster than LibreOffice and calculates on all cores.

## Critical path (backward pass)

The rules (late dates per link type, deadlines, constraint caps, total and free slack, Critical) are in
[SCHEDULING_SPEC.md](SCHEDULING_SPEC.md) section 7. How the formulas do it without a circular reference:

Each successor row works out, for each of its links (`LINKS` = 4 in `planner/tasks.py`; every link column is
generated from it), the latest its predecessor may finish; SS / SF use the predecessor's duration, an input, so
there is no circularity. Successors sit below their predecessors, so a predecessor looks only at the rows from the
one below down to `_LastSuccessorRow`. That helper reads only input columns, so it isn't volatile, and it keeps each
look-up short. Limits are stored as keys (`31-Dec-9999 - date`, `100000 - slack`) so that one `MAXIFS` per link
answers both "is there a successor" and "what is its tightest limit".

| LibreOffice, one edit then recalculate | Before the critical path | With it |
|---|---|---|
| Sample (200 rows) | 0.04 s | 0.05 s |
| 1,000 tasks | 0.42 s | ~0.5 s |

## Constraints

The rules are in [SCHEDULING_SPEC.md](SCHEDULING_SPEC.md) section 5. In the formulas, `_ConstraintCode` is MS
Project's own code, `_ConstraintStart` turns the date into a start, and two bounds clamp the forward pass:

    start = MIN(MAX(ProjectStart, _ConstraintLow, link drivers), _ConstraintHigh)

**As Late As Possible** needs the backward pass, which reads the early dates, so it can't move them without a
circular reference. Instead the early dates stay as calculated (`_Leaf...`: links, slack and `ProjectFinish` read
them) and the shown dates (`_Shown...`, rolled up into Planned Start / Finish and Start / Finish) move later by the
task's free slack.

## Views: outline, timeline scales, Roadmap

- **Outline.** Excel row groups (one level per outline level, summary rows above) give the 1 2 3 buttons on Tasks
  (the timeline is on the same rows). Grouping is sheet structure, not a formula, so it covers the rows the workbook is built with; the Show
  column (`Level <= OutlineShowLevel`) works for any row through the table's filter.
- **Timeline scales.** Each Gantt column's date is a formula of its position (`column_start` in `planner/gantt.py`:
  `GanttStart + k x days`, or `EDATE` of the month / quarter start for GanttPeriod 30 / 91), and the bar rules use
  `period_end(date)`, so months keep their real lengths and the timeline cells still hold no formulas.
- **Roadmap.** Two ranked lists (phases, milestones) and an 18-month timeline drawn by conditional formatting only.
  A phase row adds its owner (Tasks > Owner), tasks done (two COUNTIFS on `tsk_Phase`), a RAG from its finish slip
  (`rag.slip`), its next milestone (`MINIFS` of `tsk_MilestoneKey` in the phase, then one MATCH) and a `HYPERLINK`
  to its row (`ADDRESS` from the Tasks header, so it follows inserted rows);
  the month header is its only formula. It prints on one landscape page.
- Bar labels and colours from a Settings table are not built: conditional formatting can't show per-row text or take
  a colour from a cell, so both would need a formula in every timeline cell (30,000 on the sample Gantt).

## Gantt: planned = outline, current = fill

The timeline is drawn to the right of the Tasks table, on the same rows (B-1005: one plan sheet, no copy of the
task list to keep in step). It holds no formulas; each cell's look comes from conditional-format rules over six Tasks
helper columns (`_IsMilestone`, `_DoneFrom`, `_DoneTo`, `_RemFrom`, `_RemTo`, `_ShowPlan`), one formula each per
row. A spacer column keeps the table from growing into the timeline, and the header dates are located from
`gantt_Origin` by `COLUMN()`. Excel applies every rule that is true and, for each property (fill, border, font), the first one listed
wins, so a fill rule (done, slip, critical, remaining, a planned gap) and a border rule (the planned outline) combine
in one cell. The outline is four rules (closed box on a one-day plan, start, end, middle) so the box closes on the
first and last planned day. `_ShowPlan` draws it for every task until a baseline is set (`BaselineSet`), then only
for tasks that have one: scope added after the baseline shows its bar without an outline. Slip is remaining work
after the planned finish. LibreOffice applies only the first true rule per cell, so its renderings show the fills
but not the outline; Excel shows both. The committed sample is built with `--baselined` (`sample.BASELINE`, the
plan as first estimated) so the comparison shows on first open; the scheduling tests use the sample without it.

## Sprint calendar

SprintTable's Start and End follow the Settings cadence (Sprint1Start, SprintLength) until a sprint is moved: a
typed **Set End** moves that sprint's end and every later sprint follows on from it; a **Set Start** moves one
sprint's start. Start finds the nearest sprint above with a Set End using `MAXIFS` over the typed column and No, so
it never reads another row's calculated value (no OFFSET, no circular reference).

## Progress by stage

A free % Complete means something different to everyone, so a task's progress comes from its **Stage**, picked from
the Stages table on Settings (`stg_Name`, `stg_Pct`): Not Started 0%, Design 15%, In Development 40%, Test Ready
70%, In Testing 85%, Complete 100% by default. That is weighted-milestone earned value: the same stage is the same
progress on every task, and a PMO can rename the stages or change their % in one place. `_Pct` is the % the
schedule uses: a typed % Complete (now an optional override in the DETAIL section), or else the stage's %, looked up
once per row. Everything that read % Complete (remaining duration, progress roll-up, earned
value, sprint "done", the Gantt, the XML export's PercentComplete) reads `_Pct` through `_LeafProgress`. Actual
Start and Actual Finish stay typed dates; Issues asks for them when the stage calls for them and flags a stage that
disagrees with them. Older workbooks and MS Project files carry a %, so `tasks.take_stage` picks the nearest stage
and keeps the exact % as the override when it differs: an upgraded plan schedules exactly as before.

## Text inputs: Predecessors and Resources

A task's links and people are each one cell, as in MS Project: Predecessors `12, 14SS+2d` (by UID: Excel can't
rewrite typed text when rows move, so typed links can't use the row ID) and Resources `Alex Chen, JL[50%]` (names or
initials, optional units). Formulas split each cell into hidden per-link / per-person columns
(`_Pred1Text`, `_Pred1UID`, `_Pred1Type`, `_Pred1Lag` ...; `_Res2Text`, `_Res2`, `_Res2Units`, `_Res2Work` ...), with
`TRIM(MID(SUBSTITUTE(text, ",", REPT(" ", 99)), (k-1) x 99 + 1, 99))` for the k-th item: no array formulas, no
TEXTSPLIT. Each slot looks only once the one before it found something, so empty cells cost almost nothing. These
read inputs only, so they are not volatile; the scheduling and reports read the parsed columns exactly as they read
the old input columns.

The first person uses `_Resource1` / `_Units1` and the `_Leaf...` columns; people 2-4 have their own work,
time-phasing and work done columns (`tasks.PERSON_FIELDS`). Per-resource figures sum one SUMIFS per person
slot (`resources.by_resource`, `workload.work_upto`), and each slot beyond the first is wrapped in
`IF(MaxPeople >= k, ...)`: MaxPeople (Settings) is the most people on any task, so a plan with one person per task pays
for one slot only. Week and sprint boundaries are turned into working-day numbers once (`rw_CumWD`, SprintTable
`_FirstWD` / `_LastWD`) instead of inside every grid cell's SUMIFS.

## Bulk entry, look-ups and protection (B-1011, B-1017, B-1019)

- **Live resource grids.** The Resource Report's hidden cumulative-hours grid (`rw_Cum`, which feeds the weekly
  hours and utilisation) and the per-person sprint hours are the heaviest cells after the schedule itself. Each is
  wrapped in `IF(LiveGrids="No","",...)`: `IF` evaluates only the branch it takes, so with the setting off none of
  their SUMIFS run. Tasks > Over-allocated (which reads `rw_Util`) is blank meanwhile, and Status says the grids are
  off. `perf_check --timing` measures a duration edit both ways.
- **One look-up per person.** Each person in a task's Resources cell is found once on Resources (`_Res{k}Row`: by
  name, then initials, 0 when absent). The full name, work, Issues ("not on Resources", "Did you mean ...") and
  over-allocation (the grids follow Resources' row order) reuse that position. The Cost sheet matches names on its own
  rate card, so it stays independent and deletable.
- **Protected views.** Status, Roadmap and the Resource Report are protected without a password (`core.protect_view`);
  every cell in the blue input font is unlocked, so the PM RAG, comments, notes and the status date still take
  typing. Tasks isn't: Excel can't add table rows on a protected sheet, and Allow Edit Ranges doesn't work on the
  web or iPhone.

## Reports

Reports are formulas, not Power Query, so they are current after every edit, work in Excel for the web and on a
phone, and are tested in CI. They follow three rules:

- **Tasks decides, the report looks up.** A list such as "late tasks" needs a filter and a sort. Tasks does both in
  one helper column per list (`_LateKey`, `_StartSoonKey`, `_CriticalKey`, ...): blank when the task is not on the
  list, otherwise a unique sort key (the row position breaks ties). Row *k* of the report finds the *k*-th key with
  `SMALL` / `LARGE` and `MATCH`, once, in its hidden `_Task` column; the visible columns only `INDEX` by it
  (`core.ranked_table`). No `FILTER` / `SORT` (Microsoft 365 only) and no `SUMPRODUCT` per row.
- **Totals sum leaf helpers.** Progress is weighted per detail task (`_WeightDone`, `_WeightPlanned`, zero on
  summaries), so a project total is one `SUM` and a phase total is one `SUMIFS` by `tsk_Phase`.
- **One rule, one place.** RAG thresholds are settings; `planner/rag.py` turns a measure into a severity (0 / 1 / 2)
  and the worst severity into a word, for both the status rows and the phase table. A PM RAG typed next to a
  calculated RAG overrides it.

Schedule performance without money: SPI = progress / planned progress, both weighted by duration (`sch_SPI`), so
the Status RAGs work with the Cost sheet deleted.

**Leave is text, read by formulas.** Each Leave cell holds days like `16, 17(0.5)` under a month header (`Nov-26`).
`planner/leave.py` reads it with one `SUMPRODUCT` per calendar month over the constant `{1..31}` (`leave_Days`): every
day in the text gets a value in brackets and `,N(0)` is appended for each day N, so the search for `,N(` always
succeeds. That avoids `IF` / `IFERROR` on arrays, which Excel 2019 only evaluates with Ctrl+Shift+Enter. The month's
column is found by its header (so months can be added) and each person's row once per grid row, in helper cells.
Leave changes capacity only; the grids read inputs, so a task edit doesn't recalculate them.

**History without macros.** A formula can't remember last week's value, so status history works like the baseline:
`SnapshotNow`, at the bottom of the Status sheet, is one row of formulas (names only) and the PM pastes it as values
into History's `StatusHistory`, an input table the trend charts read. A macro or script may automate the copy, but the workbook never depends on it.

Weekly resource hours use the same SUMIFS spreading as the sprint load (`workload.work_between`), anchored on
`rw_LoadOrigin`; story points are done on a task's Actual Finish (`_DoneDate`), and per-sprint points are columns
of `SprintTable`, so the release burndown is one formula per column.

| LibreOffice, one edit then recalculate | Before the reports | Status + Schedule | + Resource + Agile |
|---|---|---|---|
| Sample (200 rows) | 0.05 s | 0.06 s | 0.08-0.10 s |
| 1,000 tasks | 0.54 s | 0.57-0.61 s | 0.68-0.79 s |

The weekly resource grids first cost 0.25 s per edit at 1,000 tasks (six SUMIFS per cell). A hidden grid of
cumulative hours at each week boundary (three SUMIFS, none outside the project's dates) brought that to about
0.05 s; each week is then the difference of two neighbours.

## Cost sheet (optional)

Money lives on one sheet that nothing else reads, so a PM can delete it for an audience that shouldn't see rates
(a test deletes it and checks no name or formula is left dangling). `TaskCost` has one row per Tasks row and follows
it by position (`INDEX(tsk_..., _Row)`), like the Roadmap lists. A detail task costs its people's work x day rate / hours
per day (`RateCard`, looked up by name); a summary sums its descendants (`SUM(INDEX(col, n+1):INDEX(col,
n+tsk_Descendants))`). Actual cost = work done x rate, as MS Project calculates it, so nobody types costs. Budget =
baseline work x the task's cost per hour (cost before a baseline); PV = Budget x planned %; EV = Budget x Progress;
EAC = AC + (BAC - EV) / CPI. The rate card also compares task cost with resource cost (allocated days on the project x
day rate): the gap is paid time not assigned to tasks.

## Sprint capacity

Resources ends with two grids, person by sprint. `cap_Days` = working days in the sprint within the person's Start /
End, minus their leave, x Allocation % (0 for a Cost or Material resource). `cap_Points` = days x Focus % (Settings >
Focus factor, 80%, or the person's own; 0 for advisory roles), with 1 person-day = 1 story point. The sprint table
sums the columns: Available Days, Capacity (points), and Advisory / Overhead = the difference (meetings, ceremonies,
advisory time).

## MS Project interchange

`planner/mspdi.py` (workbook → MS Project XML) and `planner/mspdi_import.py` (MS Project XML → a new workbook, through
`make_workbook(data=ProjectData(...))`) are the only converters. The command-line scripts call them, and
`docs/export.html` runs the same files in Pyodide: `scripts/build_export_page.py` bundles the planner package with
openpyxl / et_xmlfile (pinned in `requirements-dev.txt`) into the page as one deterministic zip. Rebuild the page after
changing anything in `planner/` (CI checks it); CI also drives the page in headless Chromium in both directions and
compares its output with the command line's.

## Upgrading workbooks

Users keep their plans in the workbook, and new versions move and rename inputs, so `planner/upgrade.py` moves a plan
into the latest template: it reads only typed values (formula cells are skipped), always by table and header name.
Tasks, people, assignments, holidays, leave and settings go through `ProjectData`, like an MS Project import, so the
builder sizes the tables and keeps UIDs unique; every other input table (logs, status history, sprint goals, PM RAG
and notes) is copied column by column onto the new template's input columns (blue headers) with the same header.
Dropdown lists keep the template's values and add the user's own. Each workbook carries `PlannerVersion`; a change
that moves an input adds its mapping (`RENAMED`, `RENAMED_TABLES`, `LOGS`, or a step in `read_planner`) and a fixture
of the old layout under `tests/fixtures/`, and `tests/test_upgrade.py` checks that the upgraded plan schedules exactly
as the old one did. Tables copy by name wherever their sheet now is (the Status, Agile and Log sheets of v1.7 took
over `StatusRAG`, `StatusNotes`, `SprintTable` and the rest unchanged). The three logs of v1.6 merge into the one
`Log`: RAID items keep their Ref, decisions and change requests take the next free ones, and a warning lists the new
numbers.

**Reordering** (`planner/reorder.py`, B-1015) uses the same read and rebuild. Each link becomes a constraint between
the two tasks' ancestors that are siblings under their closest common parent; each set of siblings is sorted
topologically (Kahn's algorithm, always taking the earliest original row that is free), so a task never leaves its
phase and the order changes only where a link needs it. A cycle among siblings (two phases that each need the other)
is reported, naming the links, and nothing is written.

## Project types (B-1018)

Start > Project type picks a row of `ProjectPresets` (Settings, beside the dropdown lists; editable, and a row can be
added). Sprint length and the RAG slip thresholds default to a formula that reads it, and each Stages row to
`INDEX(ProjectPresets[[Stage 1]:[Stage 6]], type, row)`: the same pattern as View from, a default anyone can type
over. The upgrade reads typed values only, so typed stages and thresholds survive and defaults stay defaults. A
typed % Complete coming in (an import or an upgrade) is matched to the type's stages.

## Modules: Simple and Full (B-913)

Agile, Log, History, Resource Report and Cost are optional (`core.MODULES`; `build_planner.py --modules` /
`--simple`). `core.set_modules` records the set for the build and `core.missing(text)` says whether a formula needs
a module that isn't there, by its name prefixes (`spr_`, `agl_`, `cap_`, `log_`, `rw_`, `evm_` ...). Each builder
then leaves out what can't calculate instead of pointing at a missing sheet: a Tasks column whose formula reads one
becomes blank (Sprint is hidden too), the sprint settings and capacity grids go without Agile, Status drops the RAG
areas, key figures, lists and snapshot fields that read a missing module, and the Start sheet drops their steps and
reference entries. The leave-by-sprint grid lives on Resources, beside the capacity it feeds, so Agile works
without the Resource Report. `ModulesTest` builds all 32 combinations and checks that no formula, name or
validation points at a missing sheet, table or name (`perf_check.missing_references`); a few are recalculated.
`PlannerModules` records the set, so an upgrade of a Simple workbook stays Simple.

## Starting from a list (B-914)

`planner/list_import.py` reads an outline, this planner's CSV or MS Project's CSV into `ProjectData`, like the XML
import, and builds the workbook the same way: levels from indentation, bullets or `1.2.3` numbers; durations in
days, weeks, hours or months; MS Project's predecessors by ID mapped to UIDs; the people named become the Resources
rows (initials that match one of them aren't added twice). Dates aren't imported: the workbook schedules from the
durations and links, so the plan starts at the CSV's earliest Start.

## One log, typed rows

The Log holds six kinds of item in one table rather than three sheets. Every row has the same core (Ref, Type, Title,
Owner, Raised, Due, Status, linked task, Update); the columns only one kind needs (probability and impact for RAID
items, options and reasons for decisions and changes, schedule and cost impact for changes) are greyed out by
conditional formatting on rows of another type, and their formulas return blank there, so the Status sheet's
`COUNTIFS` / `SUMIFS` read the whole Log by `log_Type`. "Still open" is one helper column, `_Open` (the status is
not one of Closed, Approved, Rejected, Superseded, Implemented), used by the overdue flag, the waiting days, the
top RAID list and History. The Status dropdown is a data-validation `IF` on the row's Type that picks one of three
lists, so no `INDIRECT` is needed.

## Excel version

**Excel 2019 or Microsoft 365** (Windows or Mac). The only post-2010 functions are `MINIFS` and `MAXIFS`, stored with
the `_xlfn.` prefix Excel uses, and in Release 1.0 `TEXTJOIN` (also Excel 2019). Microsoft 365-only functions
(`XLOOKUP`, `FILTER`, `LET`, `LAMBDA`, ...) are not used: a unit test enforces this, and LibreOffice, which runs the
tests, couldn't check them.

## Release 1.0 (`--release`, `release-1.0/`)

A second edition of the same code: `core.set_release` is set once per build (by `build.make_workbook`), and
`tasks.configure` then adds the Release 1.0 columns and sections to the Tasks layout. The classic workbook is built
exactly as before; `test_release.ClassicUnchangedTest` builds both in one process. What changes:

- **Inputs read through a cleaned column.** In Release 1.0 `R("Dur")`, `R("PredIn")` and `R("ResIn")` read
  `_Days`, `_Predecessors` and `_Resources` (`tasks.ALIAS`), so every scheduling formula is the same text in both
  editions and the scheduler (docs/SCHEDULING_SPEC.md) doesn't change. `tsk_Duration` points at `_Days`;
  `tsk_DurationInput` at the typed cell.
- **"Named formulas" without LAMBDA.** The brief asked for `L_PARSE_DURATION`, `L_PARSE_PRED`, `L_PARSE_RES`,
  `L_WORKING_DAYS` and `L_IS_WORKING_DAY` in Name Manager. A named formula with parameters needs `LAMBDA`
  (Microsoft 365 only, and LibreOffice 24.2, which recalculates every test, has none), so each is one helper column
  instead, which Excel calculates once per row:

  | Asked for | Is | Formula |
  |---|---|---|
  | `L_PARSE_DURATION` | `_Days` | `3`, `3d` or `2w` → working days (`VALUE` / `LEFT` / `RIGHT`); "" when unreadable |
  | `L_PARSE_PRED` | `_Predecessors` | `TEXTJOIN` of Pred1-4 / Type / Lag as `12, 14SS+2d`, else the typed text; then the v1.7 split (`_Pred{k}UID` ...) |
  | `L_PARSE_RES` | `_Resources` | `TEXTJOIN` of Res1-4 / Units as `Alex Chen, JL[50%]`, else the typed text; then the v1.7 split |
  | `L_WORKING_DAYS`, `L_IS_WORKING_DAY` | `WORKDAY.INTL` / `NETWORKDAYS.INTL` with `WorkWeek, Holidays` (`tasks.WD`) | built-in and already fast |

- **The pick list.** `_PickKey` = the row position of each detail task; `_PickList` row k = `SMALL(_PickKey, k)`'s
  UID and name, so the Pred dropdown (`=tsk_PickList`) lists the detail tasks in plan order with no gaps.
- **Leave by day.** `LeaveDays` (under the Resources table): person, month, days 1-31; `Days` is the `TEXTJOIN` of the
  ticked days as leave text. `leave.leave_days` appends it to the person's month cell (one `MATCH` on person|month
  per look-up), so the leave grids, capacity and over-allocation read both.
- **Check and TakeSnapshot** are optional modules (`core.RELEASE_MODULES`; never in Simple). Check's lists are ranked
  tables over `_IssueKey`, `_OverAllocKey` (Tasks) and `_OverdueKey` (Log); TakeSnapshot's links go to the
  defined names `snap_Row`, `tsk_BLSource` and `tsk_Baseline`.
- **TakeBaseline** steers the BL Source columns: they read Start / Finish and the typed Baseline columns, so they
  are never circular.
- **Protection and printing** (`planner/release.py`): see BACKLOG R1-10 and R1-11.

### Template 1.1 (`--release 1.1`, `release-1.1/`)

`core.set_release("1.1")`; `core.v11()` gates everything below, so the 1.0 build is unchanged
(`test_release_11.ModulesTest.test_editions_dont_leak`, `test_release.CommittedReleaseTest`).

- **Sandbox** is one `IF` in `_Days`: `IF(AND(Sandbox="Yes",[@Sandbox Duration]<>""), parse(Sandbox Duration),
  parse(Duration))` (`tasks.parse_days`). Everything reads `_Days`, so the scheduler, the reports and the MS Project
  export follow without a second engine; Status B3, the summary and Check flag it.
- **PERT**: `Expected`, `_PertVariance` and `_PertShift` per detail task; `PertP50` / `PertP80` on Settings are a
  `WORKDAY.INTL` from `ProjectFinish` by `SUMIFS` over the critical detail tasks (the shift, plus 0.84 x the square
  root of the summed variances). An approximation: the critical path can change when the estimates do.
- **Dependencies**: Log columns Expected, Task Start, Gap (days) (`NETWORKDAYS.INTL` from Expected to the task's
  start) and `_DepKey`; Status' TopDependencies is a ranked table over it.
- **Narrative** (`status.narrative`): one concatenation of names; the parts that read a module this build leaves out
  are dropped (`core.missing`).
- **Template 1.2** (`--release 1.2`, `core.v12()`): the look is `planner/theme.py`, a last pass over the built
  workbook (a colour table for fonts, fills, borders and conditional formats; fonts; tabs; gridlines), so no builder
  knows about colours twice. It runs after protection, which still finds inputs by the classic blue. The upgrade
  recognises input headers in either look (`upgrade.INPUT_FILLS`).
- **Microsoft 365 edition** (`--release m365`, `core.m365()`, `planner/m365.py`): template 1.1 where `_Days`, the
  Log's Gap, the finish variance and a working-day check call named LAMBDAs, and Check lists FILTER the tasks. A
  LAMBDA is stored as Excel stores it: `_xlfn.LAMBDA(_xlpm.x,_xlfn.LET(_xlpm.t,...))`, and called by its name; the
  lists are one-cell formulas that spill (F-1110): openpyxl can't write Excel's dynamic-array metadata, so
  `m365.spill` wraps the workbook's `save` to add `xl/metadata.xml` (the XLDAPR type), its content type and
  relationship, and `cm="1"` on each list's cell, as Excel writes a formula typed in Microsoft 365. `scripts/perf_check.py` counts array formulas and accepts `_xlpm.`
  parameters. LibreOffice 24.2 can't calculate any of it, so the build saves it without `--recalc` and
  `tests/test_m365.py` checks structure only.
- **RACI counts** (`resources.named_in`): four `COUNTIF` patterns per role (the name alone, first, last or in the
  middle of a comma list), so "Sam" never counts "Samantha".

## Guardrails

Run on every PR by `.github/workflows/ci.yml`:

| Check | Where | Fails when |
|---|---|---|
| Unit tests: schedule, baseline, links, inserting rows, blank template | `tests/test_planner.py` | any result differs from the hand-worked dates |
| No A1 references in formulas | `scripts/perf_check.py`, also a unit test | any formula contains one |
| Excel limits: balanced parentheses, length <= 8,192, nesting <= 64 | `perf_check.py`, also a unit test | any formula breaks one (LibreOffice silently repairs some; Excel would not) |
| Real row insertion and deletion keep every formula working | `RealRowEditTest` (LibreOffice via UNO) | a task can't see an inserted row, or a deleted row leaves `#REF!` |
| Volatile functions only on the allow-list (OFFSET in 11 Tasks columns) | `perf_check.py`, `perf_budget.json`, also a unit test | any other volatile use appears |
| Formula count | `perf_check.py` | above budget |
| Recalculation time: full, a duration edit near the top, the last task, a note; sample and 1,000-task stress workbook | `perf_check.py --timing` (LibreOffice) | above budget |
| Committed workbooks match the build | unit test | the `.xlsx` wasn't rebuilt after a code change |
| Tables keep their calculated-column formulas | unit test | Excel wouldn't auto-fill inserted rows |
| Same build, same bytes | `tests/test_reproducible.py` | two builds differ, or a committed workbook isn't normalised |

The results appear in the CI run summary. A budget may be raised only with a reason in the PR description.

Run the same checks locally:

```
python -m unittest discover -s tests -v
python scripts/build_planner.py /tmp/stress.xlsx --tasks 1000 --stress 33
python scripts/perf_check.py Project_Planner.xlsx --profile sample --timing
python scripts/perf_check.py /tmp/stress.xlsx --profile stress --timing
```

`--timing` needs LibreOffice and `python3-uno`.

**Parallel tests.** `scripts/run_tests.py` runs each test class in its own process, several at once. Two copies of
LibreOffice can't share a user profile, so each worker gets its own (`PLANNER_LO_PROFILE`, read by `planner.build`,
`tests/test_planner.recalc`, `tests/lo_session` and `perf_check`) and its own UNO port range (`PLANNER_LO_PORT`).
`--shard k/n` balances classes over CI machines by the times in `tests/durations.json` (`--record` refreshes it).

**Reproducible files** (`planner/reproducible.py`). `build.save` normalises every saved workbook: LibreOffice's
random chart axis ids are renumbered in order of appearance, `docProps/core.xml`'s times become a fixed date (or
`SOURCE_DATE_EPOCH`), and every zip entry gets the same timestamp. Cells, styles and charts are untouched.

**Release workflow** (`.github/workflows/release.yml`, `scripts/release_assets.py`). A `v*` tag names an edition
(`v1.0*` release-1.0, `v1.1*` release-1.1 and the M365 edition, `v1.2*` release-1.2 and its samples). The job checks
out the tag, reruns that edition's committed-workbook tests and `release_check`, and attaches the workbooks (named
with the edition), `SHA256SUMS.txt` and the CHANGELOG section. The asset script comes from the workflow's own
commit, so older tags can be published with a manual run.
