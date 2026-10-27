# Real-world fit (v1.8): the designs agreed

The owner used v1.7 in Excel and asked for changes. This page records what was agreed, so the backlog items (E10 in
[BACKLOG.md](BACKLOG.md)) can be built without re-deciding them. Every design keeps the rules: Excel 2019 functions
only, no macros or Power Query, works in Excel alone, inside the performance budgets.

## What the review found

| Feedback | Finding | Item |
|---|---|---|
| Gantt doesn't update after a change on Tasks | The formulas follow (checked in LibreOffice: a Stage change moves the Gantt row at once). The saved files don't force automatic calculation, and Excel takes its mode from the first workbook opened. | B-1001, and B-1005 removes the copy |
| No planned vs actual on the Gantt | Planned was a thin top / bottom line, almost invisible next to fills; the sample has no baseline, so planned = current. | B-1002 |
| No way to update sprints | Sprint dates came from three settings only. | B-1003 |
| Merge Tasks and Gantt | Feasible: the timeline is conditional formatting only, so it can sit beside the table. | B-1005 |
| Baseline, variance, cost, Show out of the way | Agreed **without separate sheets**: scope keeps changing, and a second sheet would be one more place to keep in step. They stay in collapsed groups; variances cost almost nothing to calculate (one subtraction each). | B-1005 |
| Sprint capacity from people | Allocation, dates, leave, holidays and a focus split. | B-1006, B-1007 |
| Agile doesn't show a sprint's tasks | Only the current sprint was listed. | B-1007 |
| Roadmap and Schedule Report not meaningful | Both are built from Tasks, but don't show it; the Schedule Report repeats Status. | B-1008, B-1009 |
| Cost doesn't make sense | Cost was hours x an hourly rate, with a manual Actual Cost. | B-1004 |

## Gantt: planned = border, current = fill

| What | Drawn as |
|---|---|
| Planned (baseline) dates | Dark border round the bar: top and bottom, left on its first day, right on its last |
| Done (Actual Start to the progress point) | Green fill |
| Remaining, inside the plan | Blue fill |
| Slip (current dates past the planned finish) | Red fill, outside the border |
| Work moved earlier | Empty bordered gap where the plan was |
| Added after the baseline (no baseline dates) | Fill without a border |

A baseline is a habit, not a sheet: set it when the plan is approved or at the start of each release / programme
increment, and again when the Log approves a change. Status reports slip since the last baseline.

## Sprint capacity (1 person-day = 1 story point)

Per person, per sprint:

| Step | Formula | Example (2-week sprint) |
|---|---|---|
| Working days | Days in the sprint within the person's Start / End, less holidays and their leave | 10 - 1 day leave = 9 |
| Available days | x **Allocation %** (share of the person on this project; replaces Max Units) | 9 x 70% = 6.3 |
| Delivery capacity | x **Focus %** (Settings > FocusFactor, default 80%; a per-person value overrides it) | 6.3 x 80% = 5.04 points |
| Ceremonies / overhead | The rest | 6.3 x 20% = 1.26 days |

Sprint capacity = the sum of everyone's delivery capacity. Advisory roles (e.g. an architect) have Focus % = 0:
no points, but their allocated days still count, shown on the sprint table's **Advisory / overhead** line, and in
cost. The sprint table compares capacity with committed and done points (velocity).

## Cost

Day Rate per person (replaces the hourly rate). Two views on Status:

- **Task cost** = assigned person-days x units x day rate (plus an optional Fixed Cost per task): what the plan's
  work costs.
- **Resource cost** = allocated days x day rate, advisory roles included: what the team costs to run.

The gap is unassigned or overhead time. Actual cost is calculated (days worked to date x day rate), not typed.

## Open points to confirm in review

- The always-visible columns of the merged plan sheet (B-1005).
- Fixed Cost per task (B-1004) and the Roadmap's meaning: phases over time (B-1009).
