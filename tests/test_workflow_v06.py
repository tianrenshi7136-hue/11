import copy
import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')

from PySide6.QtWidgets import QApplication,QStyleOptionViewItem
from PySide6.QtCore import QCoreApplication,QEvent,Qt
from els.support_editor import SupportEditor,STRUT_FIELDS,WALER_FIELDS,BRACKET_FIELDS
from els.supports import resolve_package,validate_package
from els.critical_connections import governing_welds
from els.reference_store import reference_file
from tools.verify_aa_connections_excel import sample_package


class WorkflowTests(unittest.TestCase):
 @classmethod
 def setUpClass(cls):cls.app=QApplication.instance() or QApplication([])
 def tearDown(self):
  for w in self.app.topLevelWidgets():w.close();w.deleteLater()
  QCoreApplication.sendPostedEvents(None,QEvent.Type.DeferredDelete);self.app.processEvents()
 def editor(self):
  e=SupportEditor(['152x152x37','203x203x86','102x102x23']);e.load(sample_package());return e
 def test_case_is_entered_once_and_links_follow(self):
  e=self.editor();e.case_name.setText('ULS2');p=e.collect(10,5)
  self.assertTrue(all(m['case']=='ULS2' for m in p['members']))
  self.assertEqual(p['members'][1]['source_strut_case'],'ULS2')
  self.assertEqual(p['brackets'][0]['waler_case'],'ULS2')
  self.assertTrue(e.struts.isColumnHidden(STRUT_FIELDS.index('case')))
  self.assertEqual(validate_package(p,e.sections),[])
 def test_batch_copy_displays_same_precision_without_rounding_inputs(self):
  e=self.editor();e.struts.setCurrentCell(0,0);e.duplicate('Strut',1)
  col=STRUT_FIELDS.index('anchor_kn');e.struts.item(0,col).setText('315');e.struts.item(1,col).setText('315.0')
  texts=[]
  for row in (0,1):
   opt=QStyleOptionViewItem();e.struts.itemDelegate().initStyleOption(opt,e.struts.model().index(row,col));texts.append(opt.text)
  self.assertEqual(texts,['315.00','315.00'])
  e.struts.item(0,col).setText('315.123456');self.assertEqual(e.collect(10,5)['members'][0]['anchor_kn'],315.123456)
 def test_numeric_inputs_are_editable_without_enabling_output(self):
  e=self.editor();e.include_blinding.setChecked(False)
  self.assertTrue(e.thickness.isEnabled());e.thickness.setText('250')
  for controls in e.connection_controls.values():
   controls['enabled'].setChecked(False)
   self.assertTrue(controls['leg_mm'].isEnabled());controls['leg_mm'].setText('6')
  self.assertEqual(e.collect(10,5)['blinding_thickness_mm'],250)
 def test_k_defaults_lock_and_custom_values_survive_roundtrip(self):
  e=self.editor();t=e.struts;col=lambda k:STRUT_FIELDS.index(k)
  default=t.cellWidget(0,col('use_default_k'));self.assertTrue(default.isChecked())
  self.assertEqual(t.cellWidget(0,col('lx_equals_ly')).text(),'Lx = Ly')
  for key in ('kx','ky','klt'):
   self.assertEqual(t.item(0,col(key)).text(),'1');self.assertFalse(t.item(0,col(key)).flags() & Qt.ItemFlag.ItemIsEditable)
  default.setChecked(False);t.item(0,col('kx')).setText('0.7')
  self.assertTrue(t.item(0,col('kx')).flags() & Qt.ItemFlag.ItemIsEditable)
  p=e.collect(10,5);e.load(p);self.assertEqual(e.collect(10,5)['members'][0]['kx'],.7)
  e.struts.cellWidget(0,col('use_default_k')).setChecked(True)
  self.assertEqual(resolve_package(e.collect(10,5))['members'][0]['kx'],1)
 def test_bracket_half_length_and_weld_parameters_are_separated(self):
  e=self.editor()
  for key in ('eccentricity_m','leg_mm','strength_mpa','effective_length_mm'):
   self.assertTrue(e.brackets.isColumnHidden(BRACKET_FIELDS.index(key)))
  e.brackets.item(0,BRACKET_FIELDS.index('length_m')).setText('0.8')
  e.bracket_weld_controls['leg_mm'].setText('6')
  b=e.collect(10,5)['brackets'][0];self.assertEqual(b['eccentricity_m'],.4);self.assertEqual(b['leg_mm'],6)
 def test_connection_settings_expand_without_member_ids(self):
  e=self.editor();p=e.collect(10,5);self.assertTrue(p['connection_groups'][0]['enabled'])
  self.assertNotIn('member_id',p['connection_groups'][0])
  e.struts.setCurrentCell(0,0);e.duplicate('Strut',2);p=resolve_package(e.collect(10,5))
  self.assertEqual(sum(w['kind']=='Strut-Waler' for w in p['welds']),3)
 def test_governing_is_maximum_force_not_utilization(self):
  p=resolve_package(self.editor().collect(10,5));w=copy.deepcopy(p['welds'][0]);w['member_id']='S2';p['welds'].append(w)
  fake={('S1','LC1'):(None,{'design_actions':{'n':500,'v':0}}),('S2','LC1'):(None,{'design_actions':{'n':800,'v':0}}),('W1','LC1'):(None,{'design_actions':{'n':0,'v':252}})}
  selected=governing_welds(p,fake)
  self.assertEqual(next(w for w in selected if w['kind']=='Strut-Waler')['member_id'],'S2')
 def test_imported_reference_survives_changes_or_removal_of_original(self):
  with tempfile.TemporaryDirectory() as folder:
   root=Path(folder);source=root/'original.xls';source.write_bytes(b'original reference')
   registry=root/'mapping.json';registry.write_text(json.dumps({'file':str(source)}))
   fake_book=type('Book',(),{'sheet_names':lambda self:['FSPiv Summary']})()
   with patch('xlrd.open_workbook',return_value=fake_book):managed=reference_file(registry,role='wall')
   source.write_bytes(b'changed');self.assertEqual(reference_file(registry,role='wall').read_bytes(),b'original reference')
   source.unlink();self.assertEqual(reference_file(registry,role='wall'),managed)
