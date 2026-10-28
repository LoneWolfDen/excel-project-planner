"""The optional add-ons (B-704 VBA, B-705 Office Scripts, addons/): checked by structure, since neither VBA nor Office
Scripts run outside Excel. Each file names only defined names, tables, sheets and columns the template 1.2 workbook
has, every macro is in the README, and nothing uses a cell address.

Run:  python -m unittest tests.test_addons -v
"""
import re
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))
ADDONS = ROOT / "addons"
BAS = ADDONS / "vba" / "PlannerMacros.bas"
SCRIPTS = sorted((ADDONS / "office-scripts").glob("*.ts"))


def sources():
    return [BAS] + SCRIPTS


class AddonsTest(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        from planner.build import make_workbook
        cls.wb = make_workbook(release_edition="1.2", baselined=True)

    @classmethod
    def tearDownClass(cls):
        from planner.build import make_workbook
        from test_release import small
        make_workbook(sizes=small())

    def test_files(self):
        self.assertTrue(BAS.read_text().startswith('Attribute VB_Name = "PlannerMacros"'))
        self.assertEqual({p.name for p in SCRIPTS},
                         {"setBaseline.ts", "keepBaseline.ts", "recordSnapshot.ts", "insertTask.ts", "indentTask.ts"})
        for p in SCRIPTS:
            self.assertIn("function main(workbook: ExcelScript.Workbook", p.read_text(), p.name)

    def test_every_macro_is_documented(self):
        readme = (ADDONS / "README.md").read_text()
        for macro in re.findall(r"^Public Sub (\w+)", BAS.read_text(), re.M):
            self.assertIn(f"`{macro}`", readme, macro)
        for p in SCRIPTS:
            self.assertIn(f"`{p.name}`", readme, p.name)

    def test_names_exist(self):
        wb = self.wb
        for p in sources():
            text = p.read_text()
            names = set(re.findall(r'Names\("(\w+)"\)', text)) | set(re.findall(r'getNamedItem\("(\w+)"\)', text))
            names |= {f"tsk_Baseline{k}" for k in (2, 3)} if "tsk_Baseline" in text and "which" in text else set()
            for n in names:
                self.assertTrue(n in wb.defined_names, f"{p.name}: no defined name {n}")
            tables = set(re.findall(r'ListObjects\("(\w+)"\)', text)) | set(re.findall(r'getTable\("(\w+)"\)', text))
            all_tables = {t: ws for ws in wb.worksheets for t in ws.tables}
            for t in tables:
                self.assertIn(t, all_tables, f"{p.name}: table {t}")
            for sheet in re.findall(r'Worksheets\("([^"]+)"\)', text):
                self.assertIn(sheet, wb.sheetnames, f"{p.name}: sheet {sheet}")
            headers = {c.value for c in wb["Tasks"][5] if c.value}
            for column in set(re.findall(r'ListColumns\("([^"]+)"\)', text)) | set(
                    re.findall(r'(?:getColumnByName|column)\("([^"]+)"\)', text)):
                self.assertIn(column, headers, f"{p.name}: column {column}")

    def test_no_cell_addresses(self):
        for p in sources():
            code = "\n".join(line.split("'")[0] if p.suffix == ".bas" else line.split("//")[0]
                             for line in p.read_text().splitlines())
            self.assertEqual(re.findall(r'Range\("|\$?\b[A-Z]{1,2}\$?\d+\b(?!\.)', code), [], p.name)
