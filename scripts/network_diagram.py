"""Draw the plan's links as a network diagram (B-506): a Mermaid flowchart or Graphviz DOT, critical tasks in red.

    python scripts/network_diagram.py Project_Planner.xlsx plan.mmd            # Mermaid (GitHub, mermaid.live)
    python scripts/network_diagram.py Project_Planner.xlsx plan.dot --dot      # Graphviz: dot -Tsvg plan.dot

Reads the workbook's calculated values: open and save it in Excel first (or build it with --recalc).
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from planner.exports import network  # noqa: E402
from planner.mspdi import ExportError  # noqa: E402


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("workbook")
    p.add_argument("out")
    p.add_argument("--dot", action="store_true", help="Graphviz DOT instead of Mermaid")
    a = p.parse_args(argv)
    try:
        Path(a.out).write_text(network(Path(a.workbook).read_bytes(), "dot" if a.dot else "mermaid"))
    except ExportError as e:
        raise SystemExit(str(e))
    print("saved", a.out)


if __name__ == "__main__":
    main()
