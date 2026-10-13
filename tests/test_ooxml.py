import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from els.templates import inspect_excel, sha256
from els.ooxml import audit


class InspectionTests(unittest.TestCase):
    def test_readonly_scanner_preserves_local_names_and_distinguishes_links(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "fixture.xlsx"
            # Minimal OOXML fixture, not an engineering workbook.
            with ZipFile(path, "w") as z:
                z.writestr("xl/workbook.xml", '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Calc" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Calc!$A:$AH</definedName><definedName name="old">#REF!</definedName></definedNames></workbook>')
                z.writestr("xl/_rels/workbook.xml.rels", '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>')
                z.writestr("xl/worksheets/sheet1.xml", '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>CELL("Filename",A2)&amp;"["</f><v>0</v></c><c r="B1" t="e"><f>#REF!</f><v>#REF!</v></c><c r="C1"><f>[1]Other!A1</f><v>8</v></c></row></sheetData></worksheet>')
                z.writestr("xl/externalLinks/_rels/", "")
                z.writestr("xl/externalLinks/_rels/broken.xml.rels", "")
            before = sha256(path)
            report = inspect_excel(path)
            self.assertEqual(sha256(path), before)
            sheet = report["sheets"][0]
            self.assertEqual(sheet["external_formula_count"], 1)
            self.assertEqual(sheet["broken_formula_count"], 1)
            self.assertEqual(sheet["cached_error_count"], 1)
            self.assertEqual(sheet["print_definitions"][0]["value"], "Calc!$A:$AH")
            self.assertEqual(report["broken_name_count"], 1)
            raw = audit(path)
            self.assertEqual(len(raw["malformed_parts"]), 1)
            self.assertTrue(raw["malformed_parts"][0]["part"].endswith('broken.xml.rels'))
