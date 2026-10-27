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
| Cost doesn't make sense, and rates shouldn't travel with the plan | Cost was hours x an hourly rate, with a manual Actual Cost, on Tasks and Resources. | B-1004: a separate, deletable Cost sheet |

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

## Cost: a separate sheet the plan never reads

Rates are often sensitive and not every audience needs cost, so cost lives on its own **Cost** sheet, in the same
workbook (a second, linked workbook would break when files move, and Excel for the web and iPhone don't refresh
links reliably). The dependency runs one way: the Cost sheet reads Tasks and Resources; nothing reads the Cost
sheet. Delete it before sharing the plan and nothing else changes (a test checks this). It holds:

- the **rate card**: Day Rate per person (no longer on Resources);
- **task cost** = assigned person-days x units x day rate: what the plan's work costs;
- **resource cost** = allocated days x day rate, advisory roles included: what the team costs to run; the gap
  is unassigned or overhead time;
- **actual cost**, calculated (days worked to date x day rate), not typed;
- **earned value** (BAC, PV, EV, AC, CPI, EAC). The Status sheet keeps the schedule measures.

No Fixed Cost column.

## Plan sheet columns (B-1005)

Always visible: ID, Task, Duration, Predecessors, Resources, Stage, Actual Start / Finish, Start, Finish, Sprint,
Notes, Issues. The rest (scheduling options, planned dates and variances, baseline, agile detail) in collapsed groups.

## Roadmap (B-1009): phase progress

One row per phase. Progress is **duration-weighted**, as MS Project does it: each detail task's duration x its
progress (the Stage %, or the override), summed, divided by the phase's total duration.

| Task | Duration | Stage | Progress | Duration x progress |
|---|---|---|---|---|
| A | 2 d | Complete | 100% | 2 |
| B | 10 d | In Development | 40% | 4 |
| C | 8 d | Not Started | 0% | 0 |
| **Phase** | **20 d** | | **6 / 20 = 30%** | |

A count ("1 of 3 done") weighs a 2-day task like a 10-day one and stands still while long tasks are under way, so
it is shown beside the percentage as a label: "30% · 1 of 3 tasks done". RAG comes from the phase's finish slip
against its planned finish; the next milestone is the earliest incomplete 0-day task in the phase; Owner is a new
input on phase rows.
