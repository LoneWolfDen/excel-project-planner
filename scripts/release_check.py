"""Check the Release 1.0 workbooks against docs/RELEASE_CHECKLIST.md: the items a script can check.

    python scripts/release_check.py                 # the committed workbooks in release-1.0/
    python scripts/release_check.py --edition 1.1   # ... in release-1.1/
    python scripts/release_check.py --recalc        # also recalculate the sample in LibreOffice at 25-Oct-2026

Prints one line per item (PASS / FAIL / MANUAL) and exits 1 if anything failed. The tests, the three builds and
perf_check are separate commands (CLAUDE.md "Before pushing"); MANUAL items are done in Excel
(docs/TESTING_IN_EXCEL.md, docs/TESTING_ON_IPHONE.md).
"""
import argparse
import datetime as dt
import re
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))
import openpyxl  # noqa: E402

import perf_check  # noqa: E402

FOLDERS = {"1.0": "release-1.0", "1.1": "release-1.1", "m365": "release-1.1-m365", "1.2": "release-1.2"}
LABELS = {"1.0": "1.0", "1.1": "1.1", "m365": "1.1 M365", "1.2": "1.2"}       # TemplateVersion and the README heading


def files(edition):
    folder = ROOT / FOLDERS[edition]
    return {k: folder / n for k, n in (("sample", "Project_Planner.xlsx"), ("blank", "Project_Planner_Blank.xlsx"),
                                       ("simple", "Project_Planner_Simple.xlsx"))}


FILES = files("1.0")
STATUS_DATE = dt.date(2026, 10, 25)
ERRORS = ("#NAME?", "#VALUE!", "#REF!", "#DIV/0!", "#N/A", "#NUM!", "#NULL!", "Err:")
CF_MAX = 120                       # Release 1.0's budget: fewer rules than v1.8's 121
SIZE_MAX = 5 * 1024 * 1024


def errors(wb):
    out = []
    for ws in wb.worksheets:
        charts = set()
        for table in ws.tables.values():
            cells = list(ws[table.ref])
            for i, h in enumerate(cells[0]):
                if str(h.value).startswith("_Chart"):
                    charts.update(row[i].coordinate for row in cells[1:])
        out += [f"{ws.title}!{c.coordinate}" for row in ws.iter_rows() for c in row
                if isinstance(c.value, str) and c.value.startswith(ERRORS)
                and not (c.value == "#N/A" and c.coordinate in charts)]
    return out


INPUT_BLUES = {"0000FF", "0F6CBD"}            # the input font in the classic look and in template 1.2's
OPTIONAL_SHEETS = ("Agile", "Log", "History", "Resource Report", "Cost", "Check", "TakeSnapshot")
V12_SHEETS = ("Board", "Backlog")              # template 1.2 only


def lock_problems(wb):
    """Cells whose lock flag is wrong for Review > Protect Sheet: a blue input that is locked, or a formula (not
    blue: a blue formula is a default to type over) that is not. Every sheet, not only Tasks."""
    out = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                rgb = c.font.color.rgb if c.font is not None and c.font.color is not None else None
                blue = isinstance(rgb, str) and rgb[-6:].upper() in INPUT_BLUES
                formula = isinstance(c.value, str) and c.value.startswith("=")
                if (blue and c.protection.locked) or (formula and not blue and not c.protection.locked):
                    out.append(f"{ws.title}!{c.coordinate}")
    return out


def sheet_mentions(wb, sheets):
    """Formulas, defined names, validations and chart series that name one of these sheets."""
    texts = [(f"{s}!{c.coordinate}", f) for s, c, f in perf_check.formulas(wb)]
    texts += [(f"name {n}", d.attr_text or "") for n, d in wb.defined_names.items()]
    for ws in wb.worksheets:
        texts += [(f"{ws.title} validation", f"{dv.formula1} {dv.formula2}")
                  for dv in ws.data_validations.dataValidation]
    out = []
    for where, text in texts:
        for sheet in sheets:
            if f"'{sheet}'!" in text or re.search(rf"(?<![\w']){re.escape(sheet)}!", text):
                out.append(f"{where}: {sheet}")
    return out


def validation_problems(wb):
    """Data validations with no cells, or a list / formula naming something that isn't defined."""
    names = set(wb.defined_names.keys())
    out = []
    for ws in wb.worksheets:
        for dv in ws.data_validations.dataValidation:
            if not str(dv.sqref).strip():
                out.append(f"{ws.title}: a rule with no cells")
            for f in (dv.formula1, dv.formula2):
                for word in re.findall(r"\b([A-Za-z_][A-Za-z0-9_.]*)\b(?!\()", str(f or "").lstrip("=")):
                    if "_" in word and word not in names and not re.match(r"^[A-Z]{1,3}\d+$", word):
                        out.append(f"{ws.title}: {word} in {f}")
    return out


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--recalc", action="store_true", help="recalculate the sample at 25-Oct-2026 with LibreOffice")
    p.add_argument("--edition", default="1.0", choices=list(FOLDERS),
                   help="which release folder to check (m365: saved without values, so the error check only "
                        "covers what the build wrote; Excel calculates it on open)")
    a = p.parse_args(argv)
    from planner.build import RELEASE_VIEWS
    global FILES
    FILES = files(a.edition)
    results = []

    def item(label, ok, detail=""):
        results.append((label, "PASS" if ok else "FAIL", detail))

    books = {k: openpyxl.load_workbook(v) for k, v in FILES.items()}
    values = {k: openpyxl.load_workbook(v, data_only=True) for k, v in FILES.items()}
    status = values["sample"]
    sd = status["Status"]["C5"].value
    item("Status date of the sample is 25-Oct-2026", (sd.date() if hasattr(sd, "date") else sd) == STATUS_DATE, sd)
    found = {k: errors(v) for k, v in values.items()}
    item("No formula errors in the three workbooks (as calculated by the build)", not any(found.values()),
         {k: v[:3] for k, v in found.items() if v})
    if a.recalc:
        sys.path.insert(0, str(ROOT / "tests"))
        import tempfile
        from test_planner import recalc
        with tempfile.TemporaryDirectory() as tmp:
            again = recalc(FILES["sample"], tmp)
        item("No formula errors after recalculating the sample at 25-Oct-2026", not errors(again), errors(again)[:3])
    reports = {k: perf_check.static_report(v) for k, v in FILES.items()}
    item("No A1 references in formulas", not any(r["a1_references"] for r in reports.values()))
    item("No OFFSET outside the scheduling helpers; no INDIRECT, TODAY, NOW",
         all(r["volatile"] in ({}, {"Tasks": {"OFFSET": r["volatile"].get("Tasks", {}).get("OFFSET", 0)}})
             and r["volatile"].get("Tasks", {}).get("OFFSET", 0) <= 2200 for r in reports.values()),
         {k: r["volatile"] for k, r in reports.items()})
    item("No references to a missing sheet, table or name (Simple included)",
         not any(r["missing_references"] for r in reports.values()))
    bad_dv = {k: validation_problems(v) for k, v in books.items()}
    item("Every data validation has cells and names that exist", not any(bad_dv.values()),
         {k: v[:3] for k, v in bad_dv.items() if v})
    cf = {k: sum(x["rules"] for x in r["cf"].values()) for k, r in reports.items()}
    per_row = {k: [s for s, x in r["cf"].items() if x["single_row"] > perf_check.SINGLE_ROW_RANGES_MAX]
               for k, r in reports.items()}
    item(f"Conditional formatting: at most {CF_MAX} rules, none per row", max(cf.values()) <= CF_MAX
         and not any(per_row.values()), cf)
    unprotected = [s for s in RELEASE_VIEWS if s in books["sample"].sheetnames
                   and not (books["sample"][s].protection.sheet and books["sample"][s].protection.password is None)]
    item("Report sheets protected without a password", not unprotected, unprotected)
    tasks = books["sample"]["Tasks"]
    head = {tasks.cell(5, c).value: c for c in range(1, tasks.max_column + 1)}
    locks = {k: lock_problems(v) for k, v in books.items()}
    item("Every sheet: blue inputs unlocked, formulas locked (ready for Review > Protect Sheet)",
         not any(locks.values()), {k: v[:3] for k, v in locks.items() if v})
    titled = {s: books["sample"][s].print_title_rows for s in ("Status", "Roadmap", "Resource Report")
              if s in books["sample"].sheetnames}
    item("Report pages repeat their title rows (project name, status date) when printed",
         all(t and t.startswith("$1:") for t in titled.values()), titled)
    picks = [dv.formula1.lstrip("=") for b in books.values() for dv in b["Tasks"].data_validations.dataValidation
             if dv.formula1 and "tsk_PickList" in dv.formula1]       # (LibreOffice saves it without the "=")
    item("Pred dropdowns list only the tasks above the row (a range, no OFFSET)",
         picks and all(f.startswith("INDEX(tsk_PickList,1):INDEX(tsk_PickList,") and "OFFSET" not in f
                       for f in picks), sorted(set(picks))[:2])
    if "Check" in books["sample"].sheetnames:
        item("Check counts leave typed twice (day grid and month cell)",
             any("lv_Twice" in str(c.value) for row in books["sample"]["Check"].iter_rows() for c in row))
    sizes = {k: v.stat().st_size for k, v in FILES.items()}
    item("Each file under 5 MB", max(sizes.values()) < SIZE_MAX, {k: f"{v / 1e6:.1f} MB" for k, v in sizes.items()})
    external = []
    for k, path in FILES.items():
        with zipfile.ZipFile(path) as z:
            external += [f"{k}: {n}" for n in z.namelist() if "externalLink" in n]
        external += [f"{k}: {f}" for _, _, f in perf_check.formulas(books[k]) if re.search(r"\[[^\]]+\.xls", f)]
    item("No external links", not external, external[:3])
    item("Simple: six sheets, 12 tasks", books["simple"].sheetnames == ["Start", "Status", "Roadmap", "Settings",
                                                                        "Tasks", "Resources"]
         and sum(1 for r in values["simple"]["Tasks"].iter_rows(min_row=6, min_col=head["Task Name"],
                                                                 max_col=head["Task Name"], values_only=True)
                 if r[0]) == 12)
    simple_refs = sheet_mentions(books["simple"], OPTIONAL_SHEETS)
    item("Simple: nothing names Agile, Log, History, Resource Report, Cost, Check or TakeSnapshot",
         not simple_refs, simple_refs[:3])
    if a.edition != "1.2":
        v12_refs = [f"{k}: {m}" for k, b in books.items() for m in sheet_mentions(b, V12_SHEETS)]
        v12_refs += [f"{k}: sheet {s}" for k, b in books.items() for s in V12_SHEETS if s in b.sheetnames]
        item("No template 1.2 sheet (Board, Backlog) or reference to one", not v12_refs, v12_refs[:3])
    label = LABELS[a.edition]
    item(f"TemplateVersion is {label} in every workbook",
         all(b.defined_names["TemplateVersion"].attr_text == f'"{label}"' for b in books.values()))
    readme = (ROOT / "README.md").read_text()
    item(f"README has the Release {label} and Limits sections",
         f"## Release {label}" in readme and "## Limits" in readme)
    page = (ROOT / "docs" / "export.html").read_text()
    loads = sorted(set(re.findall(r"https?://[^\"' )]+", page)))
    item("export.html loads nothing but Pyodide (the workbook never leaves the device)",
         all("cdn.jsdelivr.net/pyodide" in u for u in loads), loads)
    for label in ("Opens in Excel desktop, web and iPhone; protection and pickers work (TESTING_IN_EXCEL.md, "
                  "TESTING_ON_IPHONE.md)", "export.html: export, upgrade, import a CSV / list / XML, reorder, in a "
                  "browser"):
        results.append((label, "MANUAL", ""))

    width = max(len(r[0]) for r in results)
    for label, verdict, detail in results:
        print(f"{verdict:6}  {label.ljust(width)}  {detail if detail not in ('', None, {}, []) else ''}")
    failed = [r for r in results if r[1] == "FAIL"]
    print(f"\n{len(results) - len(failed)} of {len(results)} pass or are manual; {len(failed)} failed.")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
