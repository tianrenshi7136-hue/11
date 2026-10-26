import copy,os,unittest,tempfile
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')
from pathlib import Path
from unittest.mock import patch
from PySide6.QtWidgets import QApplication,QComboBox,QMessageBox,QPushButton
from PySide6.QtCore import QCoreApplication,QEvent,Qt
from els.support_editor import SupportEditor,STRUT_FIELDS,WALER_FIELDS,BRACKET_FIELDS
from els.generation import generation_plan
from els.ui import Workbench
from tools.verify_aa_connections_excel import sample_package


class WorkflowV07Tests(unittest.TestCase):
 @classmethod
 def setUpClass(cls):cls.app=QApplication.instance() or QApplication([])
 def tearDown(self):
  for w in self.app.topLevelWidgets():w.close();w.deleteLater()
  QCoreApplication.sendPostedEvents(None,QEvent.Type.DeferredDelete);self.app.processEvents()
 def editor(self):
  e=SupportEditor(['152x152x37','203x203x86','102x102x23']);e.load(sample_package());return e
 def test_new_waler_defaults_follow_span_and_allow_override(self):
  e=self.editor();e.add_member('Waler');t=e.walers;r=t.rowCount()-1;col=lambda k:WALER_FIELDS.index(k)
  self.assertTrue(t.cellWidget(r,col('lx_equals_ly')).isChecked())
  t.item(r,col('load_span_m')).setText('4.2')
  self.assertEqual(t.item(r,col('lx_m')).text(),'4.2');self.assertEqual(t.item(r,col('ly_m')).text(),'4.2')
  self.assertFalse(t.item(r,col('ly_m')).flags() & Qt.ItemFlag.ItemIsEditable)
  t.cellWidget(r,col('lx_equals_ly')).setChecked(False);t.item(r,col('lx_m')).setText('3');t.item(r,col('ly_m')).setText('2');t.item(r,col('load_span_m')).setText('5')
  self.assertEqual(t.item(r,col('lx_m')).text(),'3');self.assertEqual(t.item(r,col('ly_m')).text(),'2')
  p=e.collect(10,5);e.load(p);self.assertEqual(e.collect(10,5)['members'][-1]['lx_m'],3)
 def test_strut_new_row_defaults_ly_link(self):
  e=self.editor();e.add_member('Strut');self.assertTrue(e.struts.cellWidget(1,STRUT_FIELDS.index('lx_equals_ly')).isChecked())
 def test_bracket_waler_menu_tracks_defined_members_and_common_case(self):
  e=self.editor();combo=e.brackets.cellWidget(0,BRACKET_FIELDS.index('waler_id'));self.assertIsInstance(combo,QComboBox);self.assertFalse(combo.isEditable())
  self.assertEqual(combo.currentText(),'W1 / LC1');e.case_name.setText('ULS1');p=e.collect(10,5)
  self.assertEqual(p['brackets'][0]['waler_case'],'ULS1');self.assertEqual(p['brackets'][0]['waler_id'],'W1')
  e.walers.setRowCount(0);e.refresh();self.assertTrue(generation_plan({},e.collect(10,5))['missing'])
 def test_missing_modules_are_skipped_without_mutating_draft(self):
  e=self.editor();p=e.collect(10,5);p['include_blinding']=True;p['blinding_thickness_mm']=None
  original=copy.deepcopy(p);plan=generation_plan({},p)
  self.assertIn('钢板桩',plan['missing']);self.assertIn('Blinding 垫层',plan['missing']);self.assertFalse(plan['support']['include_blinding']);self.assertEqual(p,original)
  p['connection_groups'][0]['strength_mpa']=None;plan=generation_plan({},p)
  self.assertEqual(len(plan['support']['connection_groups']),1)
 def test_declining_confirmation_does_not_start_excel(self):
  with tempfile.TemporaryDirectory() as folder:
   w=Workbench(Path(folder)/'test.db');w.support_editor.load(sample_package())
   with patch.object(QMessageBox,'question',return_value=QMessageBox.StandardButton.No) as question,patch('els.ui.SupportWorker') as worker:
    w.generate_support_book();question.assert_called_once();worker.assert_not_called()
   self.assertFalse(any('验算指定型号并生成' in b.text() for b in w.findChildren(QPushButton)))
 def test_accepting_confirmation_generates_complete_modules_only(self):
  with tempfile.TemporaryDirectory() as folder:
   w=Workbench(Path(folder)/'test.db');w.support_editor.load(sample_package())
   with patch.object(QMessageBox,'question',return_value=QMessageBox.StandardButton.Yes),patch('els.ui.SupportWorker') as worker:
    w.generate_support_book();worker.assert_called_once();worker.return_value.start.assert_called_once()
    p=worker.call_args.args[0];self.assertEqual(len(p['members']),2);self.assertFalse(p['include_blinding'])
   w.support_worker=None
