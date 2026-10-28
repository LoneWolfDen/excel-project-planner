# Scheduling spec

The exact rules the Tasks sheet schedules by. Each section names the formula columns that carry it (in
`planner/tasks.py`) and the tests that check it. `tests/reference_schedule.py` is an independent Python version of
this page: `ReferenceScheduleTest` schedules the committed sample, and four edited plans that exercise every section,
from the inputs alone and compares every task's Start, Finish, planned dates, late dates, slack and Critical flag with
what the workbook calculated. If the two ever disagree, the formulas or this page is wrong.

Dates are whole days. "n working days after d" is Excel's `WORKDAY.INTL(d, n, WorkWeek, Holidays)`: d itself when
n = 0, and a fractional n is truncated.

## 1. Calendar

- Working days: the days of `WorkWeek` (Settings > Working week: Mon-Fri, Sun-Thu, Mon-Sat or Custom, a seven-letter
  mask Monday to Sunday, `1` = not working), less the dates in `HolidayTable`.
- Durations and lags are in working days. A difference between two dates (`wdiff` in `planner/tasks.py`; slack and
  variances) is the signed count of working days from the first to the second, 0 for the same day.
- `StatusWorkday` is the first working day on or after the status date.

Tests: `ScheduleTest.test_weekend_skipped`, `test_holidays_skipped`; `ReferenceScheduleTest.test_status_date_and_calendar`
(a six-day week and an added holiday).

## 2. Links

- The Predecessors cell holds up to four links, comma- (or semicolon-) separated: `UID[type][+/-lag][d]`, e.g. `12`,
  `14SS+2d`, `9FF-1`. Type is FS (default), SS, FF or SF; the lag is in working days, negative for a lead.
- A link is **ignored**, with the reason in Link Check, when it can't be read, its UID isn't on Tasks, it names the task
  itself, an empty row, a summary task, or a task **below** this one, or it sits on a summary row. A fifth link is
  ignored with a note. An ignored link never causes an error or a circular reference.
- The "above" rule (a predecessor is a detail task listed above its successor) keeps the forward and backward passes
  consistent: the forward pass reads rows above, the backward pass rows below.

Columns: `_Pred{k}Text`, `_Pred{k}UID`, `_Pred{k}Type`, `_Pred{k}Lag`, `_Pred{k}Pos`, `_Pred{k}Problem`, Link Check.
Tests: `ScheduleTest.test_predecessor_text`, `LinkCheckTest`, `ReferenceScheduleTest.test_link_types_lags_and_ignored_links`.

## 3. Outline

- Level is the outline level (at least 1). A row is a **summary** when the next row is deeper; its tasks are the
  deeper rows directly below it. Every other named row is a **detail task**.
- Summaries take no inputs that schedule: their dates, work and progress roll up (section 8).

Columns: `_Level`, `_IsSummary`, `_Descendants`, `_IsLeaf`, WBS. Tests: `ScheduleTest.test_wbs_and_outline`,
`test_summary_rollup`.

## 4. Forward pass (detail tasks, top to bottom)

Each usable link to predecessor P gives the earliest start it allows (D = this task's duration):

| Type | Earliest start |
|---|---|
| FS | (1 + lag) working days after P's finish; lag working days after it for a milestone (D = 0) |
| SS | lag working days after P's start |
| FF | lag working days after P's finish, then D - 1 working days back |
| SF | lag working days after P's start, then D - 1 working days back |

Calculated start = the first working day on or after
`MIN(MAX(ProjectStart, constraint low bound, every link's earliest start), constraint high bound)` (section 5).
Calculated finish = start + D - 1 working days (the start itself for a milestone). A **Manual** task starts on its
Constraint Date (the project start without one).

The pass runs twice: the **calculated** schedule from calculated predecessor dates (it becomes the plan until a
baseline is set), and the **forecast** from forecast predecessor dates, actuals and the status date (section 6).

Columns: `_Pred{k}PlanDriver`, `_Pred{k}FcstDriver` (OFFSET from the header, the one volatile function: see
ARCHITECTURE "Reading another row's calculated value"), `_LeafCalcStart`, `_LeafCalcFinish`, `_LeafFcstStart`,
`_LeafFcstFinish`. Tests: `ScheduleTest.test_finish_to_start`, `test_start_to_start_with_lag`,
`test_finish_to_finish_with_lag`, `test_two_predecessors_take_latest`, `test_manual_task`, `RippleTest`.

## 5. Constraints

The Constraint column holds MS Project's eight types; the code is MS Project's ConstraintType. With a Constraint Date
and no type it is Start No Earlier Than; a type that needs a date and has none is ignored (Link Check says so).
The constraint start is the date itself, or for a finish constraint (Must Finish On, Finish No Earlier / Later Than)
D - 1 working days before it.

| Constraint | Low bound | High bound |
|---|---|---|
| As Soon As Possible | – | – |
| Start / Finish No Earlier Than | constraint start | – |
| Start / Finish No Later Than | – | constraint start |
| Must Start / Finish On | constraint start | constraint start |

Constraint dates win over links (MS Project's default "Tasks will always honor their constraint dates"); a conflict
shows as negative slack, because the backward pass caps the late finish with the same high bound (section 7).

**As Late As Possible** keeps its early dates for links, slack and the project finish, and is *shown* later by its free
slack (on both the plan and the forecast), so it finishes just in time for its successors and moves nothing else. An
ALAP task that has started, or is Manual, isn't shifted.

Columns: `_ConstraintCode`, `_ConstraintStart`, `_ConstraintLow`, `_ConstraintHigh`, `_AlapShift`, `_Shown...`.
Tests: `ScheduleTest.test_start_no_earlier_than`, `ConstraintTest`, `ReferenceScheduleTest.test_constraints_and_late_as_possible`.

## 6. Actuals, progress and the status date

- Progress is the Stage's % (Settings > Stages) or a typed % Complete, which overrides it; an Actual Finish is 100%.
- Forecast start = Actual Start when entered, otherwise the forward pass; with Settings > RescheduleAfterStatus = Yes,
  unstarted work never starts before the status date.
- Forecast finish = Actual Finish when entered; for a milestone, its start; for a started task with
  RescheduleAfterStatus = Yes, the later of start + D - 1 and the remaining work
  (`ROUNDUP(D x (1 - progress))` working days counted from `StatusWorkday`); otherwise start + D - 1.

Columns: `_Pct`, `_LeafProgress`, `_LeafFcstStart`, `_LeafFcstFinish`. Tests: `ScheduleTest.test_actuals_drive_forecast`,
`RippleTest.test_status_date_moves_unstarted_work`, `StageTest`, `ReferenceScheduleTest.test_reschedule_after_status`.

## 7. Backward pass and slack (detail tasks, bottom to top)

- `ProjectFinish` = the latest forecast finish of the detail tasks.
- Each successor S limits its predecessor P's late finish, mirroring section 4 (D = S's duration):
  FS: (1 + lag) working days before S's late start (lag for a milestone S); FF: lag before S's late finish;
  SS / SF: lag before S's late start / late finish, then P's duration - 1 working days on.
- Late finish = the earliest of `ProjectFinish`, the Deadline, every successor's limit and the constraint high bound
  (as a finish: + D - 1). Late start = late finish - (D - 1) working days.
- Total Slack = working days from the forecast finish to the late finish. Free Slack = the smallest gap, in working
  days, between a link's earliest start and the successor's forecast start, capped at total slack.
- Critical = an unfinished detail task with total slack <= Settings > CriticalSlackDays. A missed deadline gives
  negative slack to the task and the chain feeding it.

Columns: `_Pred{k}FinishKey`, `_Pred{k}SlackKey`, `_LastSuccessorRow`, `_SuccessorFinishLimit`, `_LeafLateFinish`,
`_LeafLateStart`, `_LeafTotalSlack`, Free Slack, `_LeafCritical`. Tests: `ScheduleTest.test_critical_path`,
`test_deadlines`, `ReferenceScheduleTest` (every task's late dates and slack).

## 8. Roll-ups and the plan

- A summary's Start / Finish (and calculated start / finish) are the earliest start and latest finish of its detail
  tasks (as shown, after any ALAP shift); Late Start / Finish the earliest / latest; Total Slack the smallest;
  Critical when any of them is.
- Progress rolls up duration-weighted (each detail task's duration x its progress, over their total duration), as MS
  Project does it.
- Planned Start / Finish = the Baseline Start / Finish once pasted (section "Baseline" on Tasks), otherwise the
  calculated dates. Finish Var / Start Var = working days from planned to forecast.

Columns: `_CalcStart`, `_CalcFinish`, Start, Finish, Late Start, Late Finish, Total Slack, Critical, Progress, Planned
Start / Finish. Tests: `ScheduleTest.test_summary_rollup`, `BaselineTest`, `StageTest`.
