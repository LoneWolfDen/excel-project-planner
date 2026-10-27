"""History sheet: status snapshots for trend charts (Excel Table "StatusHistory").

A formula can't remember last week's value, so history is kept the way the baseline is: SnapshotNow, at the
bottom of the Status sheet, shows this status date's numbers (formulas), and on each reporting date the PM copies
that row and pastes it as values onto the first empty row of StatusHistory (inputs). The trend charts read StatusHistory: forecast vs planned finish,
earned value S-curve, SPI / CPI, and a story-point burn-up that shows scope changes. Nothing here needs a macro;
a macro or script would only automate the same copy and paste.
"""
import datetime as dt

from openpyxl.chart import LineChart, Reference
from openpyxl.utils import get_column_letter

from .core import DATE, MODULES, Col, add_table, font, title

D = dt.date
# header, formula (names only), number format
FIELDS = [
    ("Status Date", "StatusDate", DATE),
    ("Forecast Finish", "ProjectFinish", DATE),
    ("Planned Finish", "PlannedFinish", DATE),
    ("Finish Var", "FinishVarianceDays", "+0;-0;0"),
    ("% Earned", "evm_PctEarned", "0%"),
    ("% Planned", "evm_PctPlanned", "0%"),
    ("BAC", "evm_BAC", "#,##0"),
    ("PV", "evm_PV", "#,##0"),
    ("EV", "evm_EV", "#,##0"),
    ("AC", "evm_AC", "#,##0"),
    ("SPI", "evm_SPI", "0.00"),
    ("CPI", "evm_CPI", "0.00"),
    ("EAC", "evm_EAC", "#,##0"),
    ("Points Done", "agl_Done", "0.#"),
    ("Total Points", "TotalPoints", "0.#"),
    ("Open Risks", 'COUNTIFS(log_Type,"Risk",log_Open,1)', "0"),
    ("Open Issues", 'COUNTIFS(log_Type,"Issue",log_Open,1)', "0"),
    ("Pending Changes", 'COUNTIFS(log_Type,"Change",log_Open,1)', "0"),
    ("RAG", "ProjectRAG", "@"),
]
HISTORY_ROWS = 60
# fields that read a module (MODULES): left out with it
MODULE_FIELDS = {"evm": {"% Earned", "% Planned", "BAC", "PV", "EV", "AC", "SPI", "CPI", "EAC"},
                 "agile": {"Points Done", "Total Points"}, "log": {"Open Risks", "Open Issues", "Pending Changes"}}


def fields(modules=MODULES):
    """The snapshot / history fields of a workbook built with these modules."""
    drop = set().union(*(f for m, f in MODULE_FIELDS.items() if m not in modules))
    return [f for f in FIELDS if f[0] not in drop]


def sample_history():
    """Four weekly snapshots before the sample's status date (illustrative; the PM pastes the real ones)."""
    rows = []
    bac = 58040
    for when, finish, slip, earned, planned, cost_ratio, done, total, risks, issues, pending, rag in (
            (D(2026, 9, 27), D(2026, 12, 28), 0, 0.17, 0.19, 1.02, 5, 76, 2, 0, 0, "Green"),
            (D(2026, 10, 4), D(2026, 12, 29), 1, 0.27, 0.31, 1.04, 18, 78, 3, 0, 1, "Green"),
            (D(2026, 10, 11), D(2026, 12, 31), 3, 0.36, 0.43, 1.05, 23, 82, 3, 0, 0, "Amber"),
            (D(2026, 10, 18), D(2027, 1, 4), 5, 0.44, 0.54, 1.06, 33, 82, 4, 1, 0, "Amber")):
        ev, pv = round(bac * earned), round(bac * planned)
        ac = round(ev * cost_ratio)
        cpi = ev / ac
        rows.append([when, finish, D(2026, 12, 28), slip, earned, planned, bac, pv, ev, ac, round(ev / pv, 2),
                     round(cpi, 2), round(ac + (bac - ev) / cpi), done, total, risks, issues, pending, rag])
    return rows


def build_history(wb, blank, modules=MODULES):
    ws = wb.create_sheet("History")
    title(ws, "Status History",
          "On each reporting date: on the Status sheet, select the SnapshotNow row at the bottom, Copy, then click "
          "the first empty row of the table below and Paste Special > Values. The trend charts update from it.")
    ws.column_dimensions["A"].width = 3
    kept = fields(modules)
    keep = [i for i, f in enumerate(FIELDS) if f in kept]
    data = [] if blank else [[row[i] for i in keep] for row in sample_history()]
    top = 5
    ws.cell(top - 1, 2, "StatusHistory: one row per reporting date (values pasted from Status > SnapshotNow)").font = \
        font(bold=True, color="1F4E78")
    add_table(ws, "StatusHistory", top, 2, [Col(h, "in", 11 if fmt != DATE else 13, fmt, center=True,
                                                values=[row[i] for row in data])
                                            for i, (h, _, fmt) in enumerate(kept)], HISTORY_ROWS)
    _charts(ws, top, kept)
    ws.freeze_panes = f"C{top + 1}"
    return ws


def _charts(ws, top, kept):
    """Trend charts over the StatusHistory rows (charts need A1 ranges; empty rows leave no points)."""
    col = {h: 2 + i for i, (h, _, _) in enumerate(kept)}
    first, last = top, top + HISTORY_ROWS
    dates = Reference(ws, min_col=col["Status Date"], min_row=first + 1, max_row=last)

    def chart(title_text, headers, anchor, y_format=None):
        if not all(h in col for h in headers):      # a field its module left out
            return
        c = LineChart()
        c.title = title_text
        for h in headers:
            c.add_data(Reference(ws, min_col=col[h], min_row=first, max_row=last), titles_from_data=True)
        for s in c.series:
            s.smooth = False
        c.set_categories(dates)
        c.x_axis.number_format = "dd-mmm"
        if y_format:
            c.y_axis.number_format = y_format
        c.height, c.width = 7, 14
        ws.add_chart(c, anchor)

    right = 2 + len(kept) + 1
    at = lambda row: f"{get_column_letter(right)}{row}"
    chart("Forecast vs planned finish", ["Forecast Finish", "Planned Finish"], at(4), "dd-mmm-yy")
    chart("Earned value S-curve", ["PV", "EV", "AC"], at(19))
    chart("SPI and CPI", ["SPI", "CPI"], at(34), "0.00")
    chart("Story-point burn-up (done vs scope)", ["Points Done", "Total Points"], at(49))
