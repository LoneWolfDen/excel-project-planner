"""Reproducible builds (planner/reproducible.py): the same build gives the same bytes, and the committed workbooks
are stored that way, so `git status` after a rebuild shows only real changes.

Run:  python -m unittest tests.test_reproducible -v      (the recalculated build needs LibreOffice)
"""
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))
from planner.reproducible import FIXED_TIME, ZIP_TIME, normalise  # noqa: E402
from test_planner import SOFFICE  # noqa: E402

BUILD = ROOT / "scripts" / "build_planner.py"


def build(tmp, name, *args):
    out = Path(tmp) / name
    subprocess.run([sys.executable, str(BUILD), str(out), *args], check=True, capture_output=True, timeout=1800)
    return out.read_bytes()


class ReproducibleTest(unittest.TestCase):

    def test_openpyxl_build_twice(self):
        with tempfile.TemporaryDirectory() as tmp:
            a = build(tmp, "a.xlsx", "--release", "m365", "--baselined", "--simple")
            b = build(tmp, "b.xlsx", "--release", "m365", "--baselined", "--simple")
        self.assertEqual(a, b)

    @unittest.skipUnless(SOFFICE, "LibreOffice not installed")
    def test_recalculated_build_twice(self):
        """LibreOffice's random chart axis ids and save times are the parts that used to differ."""
        with tempfile.TemporaryDirectory() as tmp:
            a = build(tmp, "a.xlsx", "--release", "--baselined", "--simple", "--recalc")
            b = build(tmp, "b.xlsx", "--release", "--baselined", "--simple", "--recalc")
        self.assertEqual(a, b)

    def test_normalise_is_idempotent(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "x.xlsx"
            path.write_bytes((ROOT / "release-1.0" / "Project_Planner.xlsx").read_bytes())
            normalise(path)
            once = path.read_bytes()
            normalise(path)
            self.assertEqual(path.read_bytes(), once)

    def test_committed_workbooks_are_normalised(self):
        files = [f for f in subprocess.run(["git", "ls-files", "*.xlsx"], cwd=ROOT, capture_output=True,
                                           text=True).stdout.split() if not f.startswith("tests/fixtures/")]   # inputs, kept as found
        self.assertTrue(files)
        for f in files:
            with zipfile.ZipFile(ROOT / f) as z:
                self.assertTrue(all(i.date_time == ZIP_TIME for i in z.infolist()), f)
                core = z.read("docProps/core.xml").decode()
                self.assertIn(FIXED_TIME, core, f)
