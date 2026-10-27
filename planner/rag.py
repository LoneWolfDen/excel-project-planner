"""RAG (red / amber / green) rules shared by the Status sheet and the Schedule Report.

Thresholds are settings (RAGAmberSlipDays, RAGRedSlipDays, RAGAmberIndex, RAGRedIndex), so a PMO changes them in
one place. Each rule gives a severity (0 green, 1 amber, 2 red); several measures combine with MAX, then
rag_text turns the worst severity back into a word. Words only appear at the end, so the formulas stay short.
"""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font

from .core import fill

COLOURS = {"Green": ("C6EFCE", "006100"), "Amber": ("FFEB9C", "9C5700"), "Red": ("FFC7CE", "9C0006")}


def slip(v):
    """Finish slip in working days: amber / red from the slip thresholds."""
    return f"IF(N({v})>=RAGRedSlipDays,2,IF(N({v})>=RAGAmberSlipDays,1,0))"


def index(v):
    """SPI or CPI: amber / red below the index thresholds; no rating while there is nothing to measure."""
    return f"IF(ISNUMBER({v}),IF({v}<RAGRedIndex,2,IF({v}<RAGAmberIndex,1,0)),0)"


def count(v, severity):
    """A count that should be zero: any at all gives `severity` (1 amber, 2 red)."""
    return f"IF(N({v})>0,{severity},0)"


def rag_text(severity):
    return f'CHOOSE({severity}+1,"Green","Amber","Red")'


def colour(ws, cell_range, top_left):
    """Conditional formatting: green / amber / red fill on cells holding those words (A1, as Excel requires)."""
    for word, (bg, fg) in COLOURS.items():
        ws.conditional_formatting.add(cell_range, FormulaRule(formula=[f'{top_left}="{word}"'], fill=fill(bg),
                                                              font=Font(color=fg, bold=True)))
