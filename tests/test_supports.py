import copy,math,unittest
from els.supports import member_actions,blinding_reference,validate_package
class SupportsTests(unittest.TestCase):
 def member(self):
  return dict(member_id='S1',case='LC1',kind='Strut',section='152x152x37',grade='S355',anchor_kn_per_m=90,load_width_m=3.5,load_span_m=2.5,angle_deg=45,lx_m=2.5,ly_m=2.5,kx=1,ky=1,klt=1)
 def package(self):
  return dict(source='TEST',members=[self.member()],welds=[dict(member_id='S1',case='LC1',kind='Strut-Waler',leg_mm=8,strength_mpa=220)],length_m=10,width_m=5,blinding_thickness_mm=300)
 def test_strut_unfactored_per_metre_once_and_selfweight(self):
  a=member_actions(self.member(),47.1)
  self.assertAlmostEqual(a['n'],623.668181)
  self.assertAlmostEqual(a['v'],.6346725)
  self.assertAlmostEqual(a['mx'],.3966703125)
 def test_waler_load_span_separate_from_buckling_length(self):
  m=self.member();m.update(kind='Waler',angle_deg=0,load_span_m=4,ly_m=3.5)
  a=member_actions(m,107.8)
  self.assertAlmostEqual(a['n'],441);self.assertAlmostEqual(a['v'],252);self.assertAlmostEqual(a['mx'],252)
 def test_blinding_short_side_and_failure_boundaries(self):
  a=blinding_reference(80.4,10,5,300)
  self.assertEqual(a,blinding_reference(-80.4,5,10,300))
  self.assertAlmostEqual(a['demand_kn_per_m'],112.56)
  self.assertAlmostEqual(min(a['capacities_kn_per_m']),1160)
  self.assertEqual(a['status'],'PASS')
  self.assertEqual(blinding_reference(80.4,10,5,100)['status'],'FAIL')
  self.assertEqual(blinding_reference(1e5,10,5,300)['status'],'FAIL')
 def test_invalid_units_links_duplicates_and_missing_data(self):
  p=self.package();self.assertEqual(validate_package(p,['152x152x37']),[])
  for key,value in [('angle_deg',90),('load_width_m',None),('anchor_kn_per_m',float('nan'))]:
   x=copy.deepcopy(p);x['members'][0][key]=value;self.assertTrue(validate_package(x))
  x=copy.deepcopy(p);x['welds'][0]['kind']='Waler-Sheetpile';self.assertTrue(validate_package(x))
  x=copy.deepcopy(p);x['members'].append(self.member());self.assertTrue(validate_package(x))
