"""B-010 (template 1.2): labels that read the same in every Excel language.

TEXT(date, "dd-mmm-yyyy") depends on Excel's language: its format codes are translated (a German Excel wants
"TT.MMM.JJJJ") and the month names come out in that language. Number formats on cells are fine (Excel translates
them itself); only formulas that turn a date into text are not. This pass, run on the finished template 1.2
workbook, rewrites every TEXT(x, "<date format>") into pieces that don't depend on the language: DAY, YEAR, and the
month and weekday names spelled out with MID from constant strings, as the leave look-ups already do
(planner/leave.py). Number formats inside TEXT ("0%", "+0;-0;0") are language-independent and stay.
"""
import re

MONTHS = "JanFebMarAprMayJunJulAugSepOctNovDec"
DAYS = "MonTueWedThuFriSatSun"
PIECES = {
    "dd": lambda x: f'RIGHT("0"&DAY({x}),2)',
    "d": lambda x: f"DAY({x})",
    "mmm": lambda x: f'MID("{MONTHS}",MONTH({x})*3-2,3)',
    "ddd": lambda x: f'MID("{DAYS}",WEEKDAY({x},2)*3-2,3)',
    "yyyy": lambda x: f"YEAR({x})",
    "yy": lambda x: f"RIGHT(YEAR({x}),2)",
}
TOKEN = re.compile(r"yyyy|yy|mmm|ddd|dd|d|[ \-]")


def date_text(x, fmt):
    """The language-independent text of date x in a format made of dd / d / ddd / mmm / yy / yyyy, spaces and
    dashes; None when the format has anything else (it is then left as TEXT)."""
    parts, at = [], 0
    for m in TOKEN.finditer(fmt):
        if m.start() != at:
            return None
        tok = m.group(0)
        parts.append(f'"{tok}"' if tok in " -" else PIECES[tok](x))
        at = m.end()
    if at != len(fmt) or not parts:
        return None
    return "(" + "&".join(parts) + ")"


def _text_calls(formula):
    """(start, end, argument, format) for each TEXT(arg,"format") with a literal format."""
    for m in re.finditer(r"(?<![A-Za-z_.])TEXT\(", formula):
        i, depth, in_str, comma = m.end(), 1, False, None
        while i < len(formula) and depth:
            ch = formula[i]
            if ch == '"':
                in_str = not in_str
            elif not in_str:
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
                elif ch == "," and depth == 1 and comma is None:
                    comma = i
            i += 1
        if depth or comma is None:
            continue
        fmt = formula[comma + 1:i - 1].strip()
        if fmt.startswith('"') and fmt.endswith('"'):
            yield m.start(), i, formula[m.end():comma], fmt[1:-1]


def rewrite(formula):
    """formula with each date TEXT(...) made language-independent (innermost first, so nesting works)."""
    changed = True
    while changed:
        changed = False
        for start, end, arg, fmt in sorted(_text_calls(formula), key=lambda c: c[0], reverse=True):
            safe = date_text(arg, fmt)
            if safe:
                formula = formula[:start] + safe + formula[end:]
                changed = True
                break
    return formula


def apply(wb):
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("=") and "TEXT(" in c.value:
                    c.value = rewrite(c.value)
        for table in ws.tables.values():                 # calculated columns keep the same formula
            for column in table.tableColumns:
                f = column.calculatedColumnFormula
                if f is not None and f.attr_text and "TEXT(" in f.attr_text:
                    f.attr_text = rewrite(f.attr_text)
