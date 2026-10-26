"""Performance and structure checks for planner workbooks. Run on every PR (see .github/workflows/ci.yml).

    python scripts/perf_check.py Project_Planner.xlsx --profile sample
    python scripts/perf_check.py stress.xlsx --profile stress --timing

Static checks (openpyxl only):
  * formula count per sheet
  * volatile functions (OFFSET, INDIRECT, TODAY, NOW, RAND, ...): every use must be on the allow-list in
    perf_budget.json, because a volatile cell makes Excel recalculate it, and everything that depends on it,
    after every edit
  * A1 references inside formulas: there must be none; formulas use structured references and defined names
Timing (LibreOffice via python3-uno, with --timing):
  * full recalculation and a single-edit recalculation, median of three runs

Writes a Markdown report to stdout (and to $GITHUB_STEP_SUMMARY in CI) and exits 1 when a budget is exceeded.
"""
import argparse
import json
import os
import re
import statistics
import subprocess
import sys
import time
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl
from openpyxl.formula import Tokenizer

ROOT = Path(__file__).resolve().parents[1]
VOLATILE = ("OFFSET", "INDIRECT", "TODAY", "NOW", "RAND", "RANDBETWEEN", "CELL", "INFO")
A1 = re.compile(r"^(?:'?[^!']+'?!)?\$?(?:[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|[A-Z]{1,3}:\$?[A-Z]{1,3}|\d+:\$?\d+)$")


def formulas(wb):
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    yield ws.title, c.coordinate, c.value


def a1_references(formula, names):
    """A1-style cell or range operands in a formula (structured references and defined names are fine)."""
    found = []
    for tok in Tokenizer(formula).items:
        if tok.type == "OPERAND" and tok.subtype == "RANGE":
            v = tok.value
            if "[" in v or v in names or v.split("!")[-1] in names:
                continue
            if A1.match(v):
                found.append(v)
    return found


def static_report(path):
    wb = openpyxl.load_workbook(path)
    names = set(wb.defined_names)
    per_sheet = Counter()
    volatile = defaultdict(Counter)          # sheet -> function -> cells
    a1 = []
    for sheet, coord, f in formulas(wb):
        per_sheet[sheet] += 1
        upper = f.upper()
        for fn in VOLATILE:
            if re.search(rf"(?<![A-Z_.]){fn}\(", upper):
                volatile[sheet][fn] += 1
        refs = a1_references(f, names)
        if refs:
            a1.append(f"{sheet}!{coord}: {', '.join(refs)}")
    return {"formulas": dict(per_sheet), "formulas_total": sum(per_sheet.values()),
            "volatile": {s: dict(c) for s, c in volatile.items()},
            "volatile_cells": sum(sum(c.values()) for c in volatile.values()), "a1_references": a1}


def time_recalc(path, edit_sheet="Tasks", edit_header="Duration", runs=3):
    """Median seconds for a full recalculation and for one edit followed by an incremental recalculation."""
    import uno                                            # python3-uno, part of LibreOffice
    from com.sun.star.beans import PropertyValue

    def prop(n, v):
        p = PropertyValue()
        p.Name, p.Value = n, v
        return p

    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb[edit_sheet]
    header_row = next(r for r in range(1, 20) if any(c.value == edit_header for c in ws[r]))
    col = next(c.column_letter for c in ws[header_row] if c.value == edit_header)
    cell = f"{col}{header_row + 2}"                       # second task: an ordinary detail task in the sample
    wb.close()

    port = 2084
    proc = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore", "--nologo",
                             f"--accept=socket,host=localhost,port={port};urp;"],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        ctx = uno.getComponentContext()
        resolver = ctx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", ctx)
        for _ in range(120):
            try:
                remote = resolver.resolve(f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(0.5)
        else:
            raise RuntimeError("LibreOffice did not start")
        desktop = remote.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", remote)
        doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(str(Path(path).resolve())), "_blank", 0,
                                           (prop("Hidden", True),))
        full = []
        for _ in range(runs):
            t = time.perf_counter()
            doc.calculateAll()
            full.append(time.perf_counter() - t)
        target = doc.Sheets.getByName(edit_sheet).getCellRangeByName(cell)
        original = target.getValue()
        edit = []
        for i in range(runs):
            target.setValue(original + 1 + i)
            t = time.perf_counter()
            doc.calculate()
            edit.append(time.perf_counter() - t)
        doc.close(True)
        try:
            desktop.terminate()
        except Exception:
            pass
    finally:
        try:
            proc.wait(timeout=30)
        except subprocess.TimeoutExpired:
            proc.kill()
    return {"full_recalc_s": statistics.median(full), "edit_recalc_s": statistics.median(edit), "edited": cell}


def check(report, budget):
    problems = []
    if report["formulas_total"] > budget["formulas_max"]:
        problems.append(f"{report['formulas_total']} formulas > budget {budget['formulas_max']}")
    allowed = budget["volatile_allowed"]
    for sheet, fns in report["volatile"].items():
        for fn, n in fns.items():
            limit = allowed.get(sheet, {}).get(fn, 0)
            if n > limit:
                problems.append(f"{sheet}: {n} cells use volatile {fn} (allowed {limit})")
    if report["a1_references"]:
        problems.append(f"{len(report['a1_references'])} formulas contain A1 references, e.g. "
                        + "; ".join(report["a1_references"][:5]))
    for k in ("full_recalc_s", "edit_recalc_s"):
        if k in report and report[k] > budget[k + "_max"]:
            problems.append(f"{k} {report[k]:.3f}s > budget {budget[k + '_max']}s")
    return problems


def markdown(path, profile, report, budget, problems):
    lines = [f"### Performance: `{Path(path).name}` ({profile})", "",
             "| Metric | Value | Budget |", "|---|---|---|",
             f"| Formulas | {report['formulas_total']:,} | {budget['formulas_max']:,} |",
             f"| Volatile formula cells | {report['volatile_cells']:,} | allow-listed per sheet |",
             f"| Formulas with A1 references | {len(report['a1_references'])} | 0 |"]
    for k, label in (("full_recalc_s", "Full recalculation (LibreOffice)"), ("edit_recalc_s", "One edit, then recalculate")):
        if k in report:
            lines.append(f"| {label} | {report[k]:.3f} s | {budget[k + '_max']} s |")
    lines += ["", "Formulas by sheet: " + ", ".join(f"{s} {n:,}" for s, n in sorted(report["formulas"].items())), ""]
    vol = "; ".join(f"{s}: " + ", ".join(f"{fn} {n}" for fn, n in fns.items()) for s, fns in report["volatile"].items())
    lines += [f"Volatile functions: {vol or 'none'}", ""]
    lines += (["**Budget exceeded:**", ""] + [f"- {p}" for p in problems]) if problems else ["All budgets met."]
    return "\n".join(lines) + "\n"


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("workbook")
    p.add_argument("--profile", default="sample", help="budget profile in perf_budget.json")
    p.add_argument("--budget", default=str(ROOT / "perf_budget.json"))
    p.add_argument("--timing", action="store_true", help="also time recalculation in LibreOffice (needs python3-uno)")
    p.add_argument("--json", help="write the raw report here")
    a = p.parse_args(argv)
    budget = json.loads(Path(a.budget).read_text())[a.profile]
    report = static_report(a.workbook)
    if a.timing:
        report.update(time_recalc(a.workbook))
    problems = check(report, budget)
    md = markdown(a.workbook, a.profile, report, budget, problems)
    print(md)
    if os.environ.get("GITHUB_STEP_SUMMARY"):
        with open(os.environ["GITHUB_STEP_SUMMARY"], "a") as fh:
            fh.write(md)
    if a.json:
        Path(a.json).write_text(json.dumps(report, indent=2))
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
