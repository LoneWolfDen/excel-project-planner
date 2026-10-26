"""Build the planner workbook. See planner/build.py for options (python scripts/build_planner.py --help)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from planner.build import main  # noqa: E402

if __name__ == "__main__":
    main()
