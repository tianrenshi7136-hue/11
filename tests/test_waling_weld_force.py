import copy,unittest
from els.critical_connections import governing_welds
from els.supports import resolve_package,validate_package
from tools.verify_aa_connections_excel import sample_package


class WalingWeldForceTests(unittest.TestCase):
 def test_explicit_connection_uses_waler_force_and_strut_geometry(self):
  p=resolve_package(sample_package());sheets={('S1','LC1'):(None,{'design_actions':{'n':9899,'v':.63}}),('W1','LC1'):(None,{'design_actions':{'n':7000,'v':4000}})}
  w=next(w for w in governing_welds(p,sheets) if w['kind']=='Strut-Waler')
  self.assertEqual(w['member_id'],'S1');self.assertEqual(w['force_member_id'],'W1');self.assertEqual(w['_selection_basis'],'maximum_waling_design_shear')
 def test_several_walers_on_same_strut_select_maximum_absolute_shear(self):
  p=resolve_package(sample_package());w=copy.deepcopy(p['members'][1]);w.update(member_id='W2');p['members'].append(w)
  sheets={('S1','LC1'):(None,{'design_actions':{'n':99999,'v':.63}}),('W1','LC1'):(None,{'design_actions':{'v':400}}),('W2','LC1'):(None,{'design_actions':{'v':-700}})}
  chosen=next(w for w in governing_welds(p,sheets) if w['kind']=='Strut-Waler')
  self.assertEqual(chosen['force_member_id'],'W2');self.assertEqual(len(chosen['_candidates']),2)
 def test_missing_waler_cannot_fall_back_to_strut_axial_force(self):
  p=sample_package();p['members']=p['members'][:1];p['welds']=p['welds'][:1];p['brackets']=[]
  self.assertTrue(any('围檩' in e for e in validate_package(p)))
  with self.assertRaisesRegex(ValueError,'对应围檩'):governing_welds(p,{})
