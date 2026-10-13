import os,unittest
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt,QCoreApplication,QEvent
from els.support_editor import SupportEditor,STRUT_FIELDS,WALER_FIELDS
from els.supports import migrate_package,resolve_package,member_actions,validate_package

class EditorTests(unittest.TestCase):
 @classmethod
 def setUpClass(cls):cls.app=QApplication.instance() or QApplication([])
 def tearDown(self):
  for widget in self.app.topLevelWidgets():widget.close();widget.deleteLater()
  QCoreApplication.sendPostedEvents(None,QEvent.Type.DeferredDelete);self.app.processEvents()
 def editor(self):
  e=SupportEditor(['152x152x37','203x203x86'])
  e.add_member('Strut',dict(member_id='S1',case='LC1',section='152x152x37',grade='S355',anchor_kn=315,spacing_m=3.5,angle_deg=45,strut_length_m=2.5,lx_equals_ly=True))
  e.add_member('Waler',dict(member_id='W1',case='LC1',section='203x203x86',grade='S355',source_strut='S1 / LC1',load_span_m=4,lx_m=3.5,ly_m=3.5))
  return e
 def test_searchable_section_and_generation_button(self):
  e=self.editor();combo=e.struts.cellWidget(0,STRUT_FIELDS.index('section'))
  self.assertTrue(combo.isEditable());combo.completer().setCompletionPrefix('86');self.assertEqual(combo.completer().completionCount(),1)
  events=[];e.generate_requested.connect(lambda:events.append(True));e.generate.click();self.assertEqual(events,[True])
 def test_length_lx_ly_and_live_waler_spacing(self):
  e=self.editor();s=e.struts;col=lambda k:STRUT_FIELDS.index(k)
  self.assertEqual(s.item(0,col('lx_m')).text(),'2.5');self.assertEqual(s.item(0,col('ly_m')).text(),'2.5')
  s.item(0,col('strut_length_m')).setText('3');self.assertEqual(s.item(0,col('lx_m')).text(),'3');self.assertEqual(s.item(0,col('ly_m')).text(),'3')
  s.item(0,col('lx_m')).setText('2');s.item(0,col('strut_length_m')).setText('4');self.assertEqual(s.item(0,col('lx_m')).text(),'2')
  s.item(0,col('spacing_m')).setText('4');p=e.collect(0,0);w=p['members'][1]
  self.assertEqual(w['load_width_m'],4);self.assertAlmostEqual(w['anchor_kn_per_m'],78.75);self.assertNotIn('angle_deg',WALER_FIELDS)
 def test_bulk_copy_unique_ids_and_welds(self):
  e=self.editor();e.add_weld(dict(member_id='S1',case='LC1',kind='Strut-Waler',leg_mm=8,strength_mpa=220))
  e.struts.setCurrentCell(0,0);ids=e.duplicate('Strut',2);self.assertEqual(ids,['S2','S3']);self.assertEqual(e.welds.rowCount(),3)
  p=e.collect(0,0);self.assertEqual([m['section'] for m in p['members'][:3]],['152x152x37']*3)
 def test_normal_kn_actions_and_linked_q(self):
  e=self.editor();e.source.setText('TEST');e.add_weld(dict(member_id='S1',case='LC1',kind='Strut-Waler',leg_mm=8,strength_mpa=220))
  p=e.collect(0,0);self.assertEqual(validate_package(p,['152x152x37','203x203x86']),[])
  resolved=resolve_package(p);self.assertAlmostEqual(member_actions(resolved['members'][0],47.1)['n'],623.668181)
  w=resolved['members'][1];a=member_actions(w,107.8);self.assertAlmostEqual(a['n'],441);self.assertAlmostEqual(a['v'],252);self.assertAlmostEqual(a['mx'],252)
  p['members'][0]['spacing_m']=7;resolved=resolve_package(p)
  self.assertAlmostEqual(member_actions(resolved['members'][0],47.1)['n'],623.668181)
  self.assertAlmostEqual(resolved['members'][1]['anchor_kn_per_m'],45)
 def test_missing_source_link_rejected(self):
  e=self.editor();p=e.collect(0,0);p['members'][1]['source_strut_id']='Missing';self.assertTrue(any('关联' in x for x in validate_package(p)))
 def test_legacy_force_migration_preserves_design(self):
  old=dict(source='TEST',members=[dict(member_id='S1',case='LC1',kind='Strut',section='152x152x37',grade='S355',anchor_kn_per_m=90,load_width_m=3.5,load_span_m=2.5,angle_deg=45,lx_m=2.5,ly_m=2.5,kx=1,ky=1,klt=1)],welds=[])
  new=migrate_package(old);self.assertEqual(new['members'][0]['anchor_kn'],315);self.assertAlmostEqual(member_actions(resolve_package(new)['members'][0],47.1)['n'],member_actions(old['members'][0],47.1)['n'])
