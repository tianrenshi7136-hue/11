import os
import tempfile
import unittest
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

try:
    from PySide6.QtWidgets import QApplication
    from els.ui import Workbench
except ImportError:
    QApplication = None

from test_core import valid_project
from test_sheetpile import request as steel_request


@unittest.skipIf(QApplication is None, "PySide6 not installed")
class DesktopTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_form_save_and_preview(self):
        with tempfile.TemporaryDirectory() as folder:
            window = Workbench(Path(folder) / "test.db")
            p = valid_project()
            window.load_form(p)
            self.assertEqual(window.collect(), p)
            window.save()
            self.assertIn("已保存版本 1", window.state.text())
            window.geometry["length_m"].setValue(25)
            window.refresh_preview()
            self.assertEqual(window.preview.project["length_m"], 25)
            window.save()
            self.assertEqual(len(window.store.list_revisions()), 2)
            window.prepare()
            self.assertIn("尚未接入", window.report.toPlainText())
            window.show()
            self.app.processEvents()
            self.assertFalse(window.grab().isNull())
            window.close()

    def test_new_row_does_not_assume_zero_forces(self):
        with tempfile.TemporaryDirectory() as folder:
            window = Workbench(Path(folder) / "test.db")
            window.add_row()
            self.assertEqual(window.table.item(0, 7).text(), "")
            with self.assertRaises(ValueError):
                window.collect()
            window.close()

    def test_sheetpile_roundtrip_and_edits_expire_results(self):
        with tempfile.TemporaryDirectory() as folder:
            window = Workbench(Path(folder) / 'steel.db')
            p = valid_project()
            p['sheetpile'] = steel_request()
            window.load_form(p)
            self.assertEqual(window.collect()['sheetpile'], p['sheetpile'])
            window.save()
            self.assertEqual(window.store.load(p['id'], 1)['sheetpile'], p['sheetpile'])
            window.steel_result.setPlainText('generated')
            window.steel_fields['yield_strength_mpa'].setText('355')
            self.assertIn('已过期', window.steel_result.toPlainText())
            window.load_steel({})
            window.add_steel_case()
            self.assertIsNone(window.collect_steel()['cases'][0]['v_kn_per_m'])
            window.close()
