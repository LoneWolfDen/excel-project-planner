# Changelog

The workbooks carry their version: **TemplateVersion** (what users see, on Start and in printed footers) and
**PlannerVersion** (the code that built it, which the upgrade reads). Every edition stays in its own folder, so an
upgrade is a choice (`python scripts/upgrade_planner.py Old.xlsx New.xlsx --release <version>`). Details and the
reasons behind each item are in [docs/BACKLOG.md](docs/BACKLOG.md).

## Template 1.2 (`release-1.2/`, PlannerVersion 2.2)
- **Design:** a Fluent look (Microsoft 365 blues and greys, Segoe UI, clearer red / amber / green, banded tables),
  tab colours by role, report pages without gridlines, a landing-page Start sheet.
- **Easier inputs:** the Pred dropdown lists only the tasks above the row; dates read the same in every Excel
  language (B-010).
- **New sheets:** Board (cards by status, filtered by sprint or person; B-404) and Backlog (the product backlog with
  priorities and points; B-405), both optional and not in Simple.
- **Tasks:** MS Project's Status names (B-206), custom fields Text 1-5 / Number 1-5 / Flag 1-3 exported to MS
  Project (B-509), a one-line project summary at the top (B-111), A3 landscape printing with repeated headings (B-508).
- **Tools:** a network diagram (Mermaid or Graphviz, B-506) and CSV exports for Jira, Azure DevOps and Microsoft
  Planner (B-706); two more sample projects, a clinic extension and a partner summit (`--sample`, B-804).
- **Scheduling:** lags in % of the predecessor (`12+50%`) and in elapsed days (`12+3ed`, B-102); inactive tasks
  (Active = No, B-110); Remaining Duration and Actual Work with % Work Done (B-202); baselines 2 and 3 with a
  Compare against setting (B-201).
- **Cost:** a cost per use on the rate card and a Fixed Costs table per task (B-302).
- **Leave in one place:** the LEAVE BY DAY grid is where leave is typed; the month columns are worked out from it,
  and an upgrade moves month-cell leave onto the grid. **Predecessors by name:** Pred1-4 take a task's exact name
  too; a name two tasks share is flagged in Issues.

- **Optional add-ons** (`addons/`): VBA macros for desktop Excel (B-704) and Office Scripts for Excel for the web
  (B-705) that automate the baseline, snapshot, insert-task and indent steps, and fill in missing UIDs.

## Microsoft 365 edition (`release-1.1-m365/`, PlannerVersion 2.1-m365)
- The Check lists spill (real dynamic arrays) and the Pred dropdown lists only the tasks above (F-1110).
- A predecessor can be typed by its exact task name; Microsoft 365's dropdown also filters as you type.
- Template 1.1 with named LAMBDA functions (`L_PARSE_DURATION`, `L_WORKING_DAYS`, `L_IS_WORKING_DAY`) and FILTER
  lists on the Check sheet; a working-day check on actual dates. Saved without values; Excel calculates on open.

## Template 1.1 (`release-1.1/`, PlannerVersion 2.1)
- RACI per task, external dependencies with the five latest on Status, a Sandbox for what-if durations, an automatic
  status summary, three-point estimates with P50 / P80 finishes.
- A risk-exposure trend on History, a capacity chart on Agile, fewer conditional-formatting rules.

## Release 1.0 (`release-1.0/`, PlannerVersion 2.0), frozen
- **Ready-to-use fixes (29-Oct-2026, every release edition):** the Pred dropdown lists only the tasks above the row
  (an INDEX range, no OFFSET); leave typed on both the day grid and the month cell is shown (Also Typed Above) and
  counted on Check; report pages repeat their title rows when printed (A3 in template 1.2); `release_check` checks
  lock flags on every sheet and that Simple and 1.0 name no sheet they don't have.
- For first-time project managers: predecessor and resource pickers, Duration as 3 / 3d / 2w, leave by day, a
  duplicate UID refused, dates after the project start, the Check and TakeSnapshot sheets, TakeBaseline, protected
  report sheets, one-page printing, a 12-task Simple sample, a release checklist.

## The classic workbooks (repository root, PlannerVersion 1.8)
- **v1.8** Real-world fit: one Tasks sheet with its Gantt, planned vs current on the timeline, editable sprints, a
  deletable Cost sheet with day rates, sprint capacity per person, Status as the weekly page, the Roadmap by phase, a
  scheduling spec with a reference scheduler, protected views, project types, Simple / Full, start from a list.
- **v1.7** Simpler to use: the Start sheet, Predecessors and Resources as one cell each, one row per person with
  leave, one Status page, one Log, progress by Stage, the upgrade path.
- **v1.1 - v1.6** The core: frozen baseline, stable UIDs, the critical path and constraints, MS Project XML both
  ways, several people per task, weekly resource usage, earned value, sprints and burndown, the Roadmap.
