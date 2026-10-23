import copy
import math
import unittest

from els.connections import AA_CONNECTIONS, bracket_actions, validate_brackets
from els.supports import validate_package
from tools.verify_aa_connections_excel import sample_package


class ConnectionTests(unittest.TestCase):
    def test_point_load_units_and_single_factor(self):
        b=sample_package()['brackets'][0]
        b['eccentricity_m']=.25
        result=bracket_actions(b,107.8,29.3)
        force=1.4*77*107.8/10000*3.5
        weight=1.4*77*29.3/10000
        self.assertAlmostEqual(result['waler_force_kn'],force)
        self.assertAlmostEqual(result['shear_kn'],force+weight*.5)
        self.assertAlmostEqual(result['moment_knm'],force*.25+weight*.5**2/2)
        b['eccentricity_m']=.5
        self.assertAlmostEqual(bracket_actions(b,107.8,29.3)['moment_knm'],force*.5+weight*.5**2/2)

    def test_distributed_waler_weight_preserves_total_force(self):
        p=sample_package();b=p['brackets'][0]
        b['eccentricity_m']=.25
        point=bracket_actions(b,107.8,29.3)
        b['load_model']='uniform';b['eccentricity_m']=None
        uniform=bracket_actions(b,107.8,29.3)
        self.assertAlmostEqual(uniform['shear_kn'],point['shear_kn'])
        self.assertAlmostEqual(uniform['moment_knm'],point['moment_knm'])
        b['length_m']=1
        longer=bracket_actions(b,107.8,29.3)
        self.assertAlmostEqual(longer['waler_force_kn'],uniform['waler_force_kn'])
        self.assertGreater(longer['moment_knm'],uniform['moment_knm'])

    def test_bracket_missing_inputs_and_bad_links_block_generation(self):
        p=sample_package();library=['152x152x37','203x203x86','102x102x23']
        self.assertEqual(validate_package(p,library),[])
        for field,value in [('length_m',0),('tributary_length_m',None),('effective_length_mm',math.nan),
                            ('waler_id','S1'),('load_model',''),('eccentricity_m',.6),('grade','S999')]:
            bad=copy.deepcopy(p);bad['brackets'][0][field]=value
            self.assertTrue(validate_package(bad,library),(field,value))
        bad=copy.deepcopy(p);bad['brackets'].append(copy.deepcopy(bad['brackets'][0]))
        self.assertTrue(validate_brackets(bad))
        p['connection_template']='unsupported'
        self.assertTrue(validate_package(p,library))

    def test_ui_roundtrip_preserves_explicit_bracket_geometry(self):
        from els.support_editor import SupportEditor
        from PySide6.QtWidgets import QApplication
        from PySide6.QtCore import QCoreApplication,QEvent
        app=QApplication.instance() or QApplication([])
        editor=SupportEditor(['152x152x37','203x203x86','102x102x23'])
        editor.load(sample_package());package=editor.collect(10,5)
        self.assertEqual(package['connection_template'],AA_CONNECTIONS)
        self.assertEqual(package['brackets'],sample_package()['brackets'])
        editor.add_bracket();self.assertEqual(editor.text(editor.brackets,1,0),'B2')
        self.assertIsNone(editor.collect(10,5)['brackets'][1]['length_m'])
        editor.close();editor.deleteLater();QCoreApplication.sendPostedEvents(None,QEvent.Type.DeferredDelete);app.processEvents()
