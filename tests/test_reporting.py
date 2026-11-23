import csv
import pathlib
import subprocess
import sys
import tempfile
import unittest
from openpyxl import load_workbook

ROOT = pathlib.Path(__file__).resolve().parents[1]
EXTENDED = True

class ReportingTest(unittest.TestCase):
    def test_duplicate_and_missing_name_workbook(self):
        with tempfile.TemporaryDirectory() as directory:
            work = pathlib.Path(directory)
            with (work / "sap_lieferanten.csv").open("w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["Lieferantennr", "Name", "Land", "Telefon"])
                writer.writerows([
                    ["DEMO001", "Example Normal Supplier", "Deutschland", "012345"],
                    ["DEMO001", "Duplicate Must Be Removed", "Deutschland", "012345"],
                    ["DEMO002", "", "Deutschland", ""],
                    ["DEMO003", "Demo Restriction Alpha", "Frankreich", "+3312345"],
                ])
            (work / "eu_sanctions.csv").write_text("Firmenname\nDemo Restriction Alpha\n", encoding="utf-8")
            result = subprocess.run([sys.executable, str(ROOT / "sap_export_tool.py")], cwd=work, capture_output=True, text=True)
            self.assertEqual(result.returncode, 0, result.stderr)
            reports = list(work.glob("Lieferanten_Report_*.xlsx"))
            self.assertEqual(len(reports), 1)
            book = load_workbook(reports[0])
            all_rows = book["Alle Lieferanten"]
            self.assertEqual(all_rows.max_row, 4, "Three distinct suppliers, with no phantom formatting rows")
            self.assertEqual([all_rows.cell(i, 1).value for i in range(2, 5)], ["DEMO001", "DEMO002", "DEMO003"])
            self.assertEqual(book["Pruefliste"].max_row, 2)
            self.assertEqual(book["Pruefliste"]["A2"].value, "DEMO002")
            if EXTENDED:
                self.assertEqual(all_rows["A3"].fill.fgColor.rgb[-6:], "FFFF00")
                self.assertEqual(all_rows["A4"].fill.fgColor.rgb[-6:], "FF6B6B")
                self.assertEqual(book["SANKTIONSLISTE"]["A2"].value, "DEMO003")
                dashboard = {row[0].value: row[1].value for row in book["DASHBOARD"].iter_rows() if len(row) >= 2}
                self.assertEqual(dashboard["Ohne Name/Land"], 1)
            book.close()

if __name__ == "__main__":
    unittest.main()
