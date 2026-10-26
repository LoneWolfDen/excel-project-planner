# How close is the planner to Microsoft Project?

> **Update (v1.1):** frozen baseline, stable UIDs for links, invalid-link detection, deadlines, the full variance set,
> a blank template and CI are now in. Rows below are updated; struck-through gaps are closed.

Assessment of `Project_Planner.xlsx` as first built in commit `b7830cf`, compared against Microsoft Project desktop
(Standard / Professional) for a **single project**. Percentages are judgement estimates of feature coverage,
not measurements.

## Summary

| | Coverage |
|---|---|
| Your original requirements (see below) | **~80%** – all present, three need hardening |
| MS Project single-project feature set | **~40%** |
| MS Project including portfolio / server features | **~25%** |

The planner already covers what most people use MS Project for day to day: outline, durations,
FS/SS/FF/SF links with lag, auto/manual tasks, a working calendar, summary roll-ups, actuals that push
the schedule, a tracking-style Gantt, resources with rates, and an import path into MS Project.

The biggest gaps against MS Project are:

1. ~~**No frozen baseline.**~~ *Done in v1.1.* "Planned" dates are recalculated from durations and links, so editing a duration
   changes the plan too. MS Project's Baseline is a snapshot that never moves.
2. ~~**No critical path or slack.**~~ *Done in v1.2.*
3. ~~**Links break when rows are inserted.**~~ *Done in v1.1 (UID).* Predecessors refer to the row-sequence ID; MS Project re-points
   links automatically.
4. **One resource per task, at most two predecessors per task.**
5. ~~**Import is via the Excel Import Wizard only.**~~ *Done in v1.4:* Microsoft Project XML export (browser page and CLI).
6. **Never opened in real Excel or MS Project.** Everything was verified in LibreOffice only.

## Your requirements vs what exists

| Requirement | Status | Notes |
|---|---|---|
| Excel-based, fully compatible with MS Project | Partial | Import Wizard sheets with MS Project field names. Untested in MS Project; no XML export. |
| Tasks and sub-tasks | Done | Outline levels 1–9, auto WBS, summary roll-up. |
| Planned and actual start/end dates | Partial | Actuals are inputs. Planned dates are calculated, not frozen (see gap 1). |
| Notes | Done | Exported to MS Project's Notes field. |
| Days with formulas updating all future dependencies | Done | Forecast dates ripple through FS/SS/FF/SF links, lags and status-date rescheduling; covered by tests. |
| Gantt showing planned and actual | Done | Planned line plus done / remaining / late bars, milestones, status-date line, non-working days. |
| Resource planning | Partial | Resource sheet, rates, hours and cost, per-sprint load. One resource per task; no resource calendars. |
| Sprint planning | Partial | Generated sprint calendar, capacity vs load, per-person utilisation. No story points, burndown or manual sprint assignment. |
| Everything configurable | Mostly | Calendar, hours, scale, sprints, rescheduling are settings. Colours, row counts and Gantt length are fixed in the build. |
| Named references instead of cell references | Mostly | All cross-sheet lookups and settings use names. Same-row references (e.g. `$F6`) remain; Excel Tables would make these `[@Duration]`. |
| As consistent as MS Project | Partial | Terminology differs in places ("Planned" vs "Baseline", "Forecast" vs "Start/Finish"). |

## Feature-by-feature comparison

Legend: ✅ equivalent · 🟡 partial / simplified · ❌ missing · ⛔ not practical with formulas alone

### Scheduling engine (~45%)

| MS Project feature | Planner | Notes |
|---|---|---|
| Auto / manual scheduled tasks | ✅ | Manual pins to the date entered. |
| Outline levels, summary tasks, WBS codes | ✅ | Summary dates, work, cost, % complete roll up. |
| Project summary task (ID 0) | ❌ | Dashboard shows the equivalent numbers. |
| Milestones | 🟡 | Duration 0 only; no "Mark as milestone" for non-zero tasks. |
| Link types FS / SS / FF / SF with lag and lead | ✅ | Lag in working days. |
| Unlimited predecessors | 🟡 | Two per task. |
| Lag as % or elapsed days (`ed`) | ❌ | |
| Links survive row insert / delete (Unique ID) | ✅ | v1.1: predecessors reference a permanent UID. |
| Constraints: ASAP, SNET | ✅ | |
| Constraints: ALAP, SNLT, FNET, FNLT, MSO, MFO | ❌ | MSO ≈ Manual mode. ALAP needs a backward pass. |
| Deadlines | ✅ | v1.1: Deadline column, Deadline Check, ▼ on Gantt, exported to MSP. |
| Critical path, total / free slack | ✅ | Backward pass with deadlines; critical bars on the Gantt. |
| Task types (Fixed Units / Work / Duration), effort-driven | ❌ | Work = Duration × Hours/day × Units only. |
| Duration units (minutes, hours, weeks, months, elapsed, estimated `?`) | ❌ | Working days only. |
| Project calendar with exceptions | ✅ | Working week mask + holiday list. |
| Resource and task calendars | ❌ | |
| Split tasks, recurring tasks, inactive tasks | ❌ | |
| Status date and "reschedule incomplete work" | ✅ | Settings switch. |
| Resource levelling | ⛔ | Needs code (VBA or Python). |

### Tracking and progress (~50%)

| MS Project feature | Planner | Notes |
|---|---|---|
| Actual start / finish | ✅ | |
| % complete, roll-up to summaries | ✅ | Duration-weighted. |
| Actual / remaining duration, % work complete, physical % complete | ❌ | |
| Actual work, actual cost entry | 🟡 | v1.3: Actual Cost per task (blank = Cost x Progress). No actual work. |
| Baseline (frozen) and Baseline 1–10 | 🟡 | v1.1: one frozen baseline (start, finish, duration, work, cost) set by copy → paste values. No Baseline 1–10. |
| Variance: finish | ✅ | Working days. |
| Variance: start, duration, work, cost | ✅ | v1.1. |
| Earned value (PV, EV, AC, SPI, CPI, EAC) | ✅ | v1.3: per task and project, plus phase SPI / CPI; Status Report. |
| Tracking Gantt | ✅ | |
| Progress lines | ❌ | |
| Status indicators | 🟡 | Text status with colours. |

### Resources and cost (~30%)

| MS Project feature | Planner | Notes |
|---|---|---|
| Resource sheet (name, initials, group, max units, std rate, email) | ✅ | |
| Work / Material / Cost resource types | 🟡 | Type stored; Material and Cost resources are treated like Work. |
| Overtime rate, cost per use, cost rate tables, accrual | ❌ | |
| Fixed cost per task | ❌ | |
| Multiple resources per task (assignments) | ❌ | One resource with units. |
| Resource availability over time, leave | 🟡 | v1.4: Leave sheet (full / part days) lowers capacity; it doesn't move tasks (no resource calendars). |
| Resource Usage / Task Usage time-phased views | 🟡 | v1.3: resource x week hours and utilisation (working days only), plus per sprint. No task usage view. |
| Over-allocation detection | ✅ | Per resource per week and per sprint, counted on the Resource Report. |
| Resource graph | 🟡 | Sprint capacity vs work chart. |
| Team Planner | ❌ | |

### Views and usability (~35%)

| MS Project feature | Planner | Notes |
|---|---|---|
| Gantt Chart / Tracking Gantt | ✅ | 150 columns, daily or weekly. |
| Timescale zoom (days, weeks, months, quarters) | 🟡 | Days or weeks. |
| Critical path / deadline / baseline formatting on bars | ❌ | |
| Bar text (resource names beside bars) | ❌ | |
| Expand / collapse outline | ❌ | AutoFilter only. |
| Indented task names | 🟡 | Indented on Gantt, not on Tasks. |
| Filters, groups, highlight filters | 🟡 | AutoFilter. |
| Custom fields (Text1–30, Number, Flag) | ❌ | Notes only. |
| Network Diagram, Calendar view, Timeline view | ❌ | |
| Task Inspector / change highlighting / undo | ⛔ | |

### Reports (~30%)

| MS Project feature | Planner | Notes |
|---|---|---|
| Dashboard / Project Overview | ✅ | KPIs, status counts, next milestones, sprint chart. |
| Burndown / burnup | ✅ | v1.3: burndown from the plan; burn-up and trends from History snapshots. |
| Risk / issue / decision / change logs | ✅ | v1.3: RAID, Decisions and Changes sheets (not in MS Project desktop). |
| Cost overview, Work overview, Resource overview | 🟡 | v1.3: earned value, phase cost and the Resource Report; no time-phased cost. |
| Late tasks, upcoming tasks, milestone report | ✅ | v1.3: Schedule Report (late, starting / due in a look-ahead window, milestones, critical). |
| Project status with RAG | ✅ | v1.3: calculated RAG by area with thresholds, PM override, commentary. |

### Agile (MS Project Online desktop / Project for the web) (~30%)

| Feature | Planner | Notes |
|---|---|---|
| Sprints with dates | ✅ | Generated from start, length, count. |
| Sprint assigned to task | ✅ | v1.3: automatic from forecast start, or Planned Sprint. |
| Story points, velocity | ✅ | v1.3: per sprint, rolling velocity, forecast finish sprint. |
| Sprint / task board (Kanban) | ❌ | |
| Sprint burndown | ✅ | v1.3: release burndown by sprint and the current sprint by day. |
| Backlog | 🟡 | Items outside sprint range show "Backlog". |

### Interoperability (~40%)

| Feature | Planner | Notes |
|---|---|---|
| Import into MS Project via Import Wizard | 🟡 | Field names match; untested; blank rows import as blank tasks. |
| Microsoft Project XML (MSPDI) export | ✅ | v1.4: browser page and CLI; calendar, holidays, links by UID, constraints, baseline, resources, assignments. Element order checked against files MS Project wrote; not yet opened in MS Project. |
| Import from MS Project (.mpp / .xml) | ✅ | v1.4: MS Project XML into a new workbook (browser page and CLI); save .mpp as XML first. |
| Saved import/export maps | ❌ | |
| Excel for the web / Mac / Google Sheets | ? | Untested. OFFSET is supported everywhere; Google Sheets unknown. |

## Known limitations of the current build

These are defects or rough edges in what exists, not missing features.

1. **Untested in Excel and MS Project.** Specifically: whether `WORKDAY.INTL` / `NETWORKDAYS.INTL` written without
   the `_xlfn.` prefix open cleanly in Excel, whether the OFFSET look-ups raise circular-reference warnings in Excel,
   and whether the Import Wizard maps the `MSP_Tasks` columns as expected.
2. ~~**Planned dates move**~~ – fixed once a baseline is set (v1.1).
3. ~~**Inserting a row renumbers IDs**~~ – links use UIDs (v1.1). Insert with *Insert Copied Cells* so the new row has formulas.
4. ~~**Bad links cause circular references**~~ – self, own-summary, empty-row and missing-UID links are ignored and explained (v1.1).
   A loop through several tasks (A → B → A) is still circular.
5. ~~**Sprint and resource hours are prorated by calendar day**~~ – working days only since v1.2.
6. **Material and Cost resources** produce hours and count toward cost as if they were people.
7. ~~**Dashboard "planned progress"** only counts tasks whose planned finish has passed~~ – partial credit by working day (v1.3).
8. **Grid sizes are set at build time:** tables grow, but the Gantt timeline, per-resource grids and Resource Report weeks are
   sized by `scripts/build_planner.py` options (`--tasks`, `--resources`, `--grid-resources`, `--sprints`, `--gantt-columns`, `--weeks`).
9. **OFFSET is volatile,** so every edit recalculates the whole workbook. Fine at 200 rows; slower at 2,000+.
10. **Gantt glyphs** (`━`, `◆`, `◇`) depend on the font; they may render as dashes in some fonts.
11. **Locale:** `TEXT(date,"ddd")` / `"mmm"` labels are locale-dependent.

## Constraints of a formula-only workbook

Some MS Project behaviour cannot be done with formulas alone. Recommended split:

| Approach | Good for | Examples |
|---|---|---|
| **Formulas (.xlsx)** – keep as the core | Anything that recalculates from inputs | Scheduling, roll-ups, critical path (with a row-order rule), earned value, reports, boards |
| **Python CLI in this repo** | Conversion, generation, heavy algorithms; testable in CI | MSP XML export/import, resource levelling, network diagram, regenerating with more rows |
| **Optional VBA add-on (.xlsm)** | One-click actions inside Excel | Set Baseline, insert task keeping links, indent/outdent, expand/collapse |

The critical path needs care: a backward pass looks at successors, and the naive formula is circular in Excel.
It can be done without circularity if each task only looks at successors **below** it (the usual MS Project
ordering), with a warning flag when a successor sits above its predecessor.

## Similar tools worth comparing

| Tool | Why look at it |
|---|---|
| [ProjectLibre](https://www.projectlibre.com/) | Free desktop MS Project clone; opens .mpp; good reference for behaviour. |
| [GanttProject](https://www.ganttproject.biz/) | Free, open source; MS Project import/export, resource load charts. |
| [MPXJ](https://www.mpxj.org/) | Open-source library (Java/.NET/Python) that reads and writes .mpp, MSPDI XML and more – the natural engine for B-701/B-702. |
| Vertex42 Gantt chart templates | Popular Excel Gantt templates; useful UX benchmark for formula-only planners. |
| Microsoft Planner (Premium) / Project for the web | Microsoft's current cloud direction; sprints, boards, dependencies. |
| Smartsheet | Spreadsheet-style project management with Gantt and dependencies. |
