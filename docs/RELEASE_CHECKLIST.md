# Release checklist (Release 1.0)

Run before tagging a release of the workbooks in `release-1.0/`. `python scripts/release_check.py` checks every item
marked **script** and prints PASS / FAIL; `--recalc` also recalculates the sample in LibreOffice at 25-Oct-2026.
Items marked **test** are covered by `python -m unittest discover -s tests -v` (the test named). **Manual** items
are done in Excel with [TESTING_IN_EXCEL.md](TESTING_IN_EXCEL.md) and [TESTING_ON_IPHONE.md](TESTING_ON_IPHONE.md).

Last run: 28-Oct-2026, on the commit that added Release 1.0 (results in the PR). Template 1.1 (`release-1.1/`): the
same list with `python scripts/release_check.py --edition 1.1` and `tests/test_release_11.py`, plus a manual Sandbox
check in Excel (Settings > Sandbox = Yes: Status shows the red banner; set it back to No).

## Builds

| | Item | How |
|---|---|---|
| ☑ | The three classic workbooks rebuild unchanged (`Project_Planner.xlsx`, `_Blank`, `_Simple` at the root) | test: `CommittedWorkbooksTest` |
| ☑ | The three Release 1.0 workbooks match the build (`release-1.0/`) | test: `test_release.CommittedReleaseTest` |
| ☑ | All tests pass | `python -m unittest discover -s tests -v` |
| ☑ | perf_check passes on the classic sample, the Release 1.0 sample and the 1,000-task stress workbook | `scripts/perf_check.py ... --timing` (numbers in the PR) |

## Formulas

| | Item | How |
|---|---|---|
| ☑ | No formula errors in any of the three workbooks, at the status date 25-Oct-2026 | script; test: `CommittedSampleTest.test_no_errors` |
| ☑ | No A1 references in formulas | script; test: `StructureTest.test_formula_rules` |
| ☑ | No OFFSET outside the 2,200 scheduling helper cells; no INDIRECT, TODAY, NOW | script; test: `StructureTest.test_formula_rules` |
| ☑ | Excel 2019 functions only (MINIFS, MAXIFS, TEXTJOIN; no LET, LAMBDA, FILTER, XLOOKUP) | test: `StructureTest.test_functions_need_at_most_excel_2019` |
| ☑ | No reference to a missing sheet, table or name, Simple and every module mix included | script; test: `ModulesTest` |
| ☑ | The sample schedules the same as the classic sample and as the Python reference scheduler | test: `CommittedSampleTest` |

## Inputs and rules

| | Item | How |
|---|---|---|
| ☑ | Every data validation has cells and refers only to names that exist | script |
| ☑ | Duration takes 1, 1d, 1w, 0 and refuses "2days"; dates must be on or after the project start; a duplicate UID is refused | test: `StructureTest.test_validation`, `EditTest` |
| ☑ | Every picker shows an input message | test: `StructureTest.test_validation` |
| ☑ | Conditional formatting: at most 120 rules (v1.8: 121), none per row | script; test: `StructureTest.test_formula_rules` |

## Protection, printing, files

| | Item | How |
|---|---|---|
| ☑ | Start, Status, Roadmap, Resource Report, Agile, Log, Cost, Check and TakeSnapshot protected without a password; their blue cells unlocked | script; test: `StructureTest.test_protection` |
| ☑ | Tasks, Resources, Settings and History unprotected (rows are added there) with blue cells unlocked and formulas locked | script |
| ☐ | Protection works in Excel: a grey cell refuses typing, a blue one takes it, Review > Unprotect Sheet needs no password | manual: TESTING_IN_EXCEL.md section 4 |
| ☑ | Status, Roadmap and Resource Report print on one page with the template version in the footer | test: `StructureTest.test_printing`; manual: Print preview |
| ☑ | Each file under 5 MB, no external links | script |
| ☑ | TemplateVersion = 1.0 in every workbook | script |

## The converter page (docs/export.html)

| | Item | How |
|---|---|---|
| ☑ | It loads nothing but Pyodide; the workbook never leaves the device | script |
| ☐ | Works offline. **Known limit:** the first load needs the internet for Pyodide (the browser caches it afterwards) | manual |
| ☑ | Export to MS Project XML, upgrade, import a CSV / outline / XML and reorder, from Python | tests: `test_mspdi`, `test_upgrade`, `test_list_import`, `test_reorder`, `test_release.ToolsTest` |
| ☐ | The same four in a browser | test: `test_mspdi.ExportPageTest` (CI, headless Chromium); manual on a phone |

## Docs

| | Item | How |
|---|---|---|
| ☑ | README: Release 1.0 and Limits sections | script |
| ☑ | ARCHITECTURE: the named formulas and why there is no LAMBDA | review |
| ☑ | TESTING_ON_IPHONE: the Release 1.0 steps | review |
| ☑ | BACKLOG: Release 1.0 marked DONE; the rest listed as v1.1 Future | review |
