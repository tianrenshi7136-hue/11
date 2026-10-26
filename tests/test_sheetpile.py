import unittest
from els.sheetpile import draft, envelope, validate_request, rejection_message, generate


def request():
    r = draft()
    r.update(section='FSP IV', yield_strength_mpa=275, shear_reduction=.7,
             bending_reduction=.7, stiffness_reduction=.5, analysis_ei_knm2_per_m=39565,
             source='TEST FREW', parameter_reference='TEST PARAMETERS',
             cases=[{'name': 'A', 'v_kn_per_m': -80.4, 'm_knm_per_m': 20},
                    {'name': 'B', 'v_kn_per_m': 5, 'm_knm_per_m': -74.1}])
    return r


class SheetpileTests(unittest.TestCase):
    def test_english_headers_required_before_creating_outputs(self):
        with self.assertRaisesRegex(ValueError, '不得含中文'):
            generate(request(), 'not-read.json', 'not-created', '中文项目', 'Subject')
        with self.assertRaisesRegex(ValueError, 'Subject Name'):
            generate(request(), 'not-read.json', 'not-created', 'Project', '')
    def test_preselection_has_no_analysis_baseline_and_cannot_generate(self):
        r = request()
        r.update(purpose='preselection', selection_mode='lightest', analysis_ei_knm2_per_m=None)
        self.assertEqual(validate_request(r), [])
        with self.assertRaisesRegex(ValueError, '只验证指定型号'):
            generate(r, 'not-read.json', 'not-created', 'test')
    def test_rejection_identifies_manual_scope_and_controlling_check(self):
        msg = rejection_message([{'section': 'FSP IV', 'design_shear': 140, 'shear_capacity': 665.38,
            'shear_utilization': .2104, 'design_moment': 560, 'moment_capacity': 436.975,
            'moment_utilization': 1.281537845}], 'manual')
        self.assertIn('未自动尝试其他型号', msg)
        self.assertIn('560.00 / 436.98', msg)
        self.assertIn('128.2%，未通过', msg)
        self.assertIn('21.0%，通过', msg)
    def test_missing_input_does_not_default_to_template(self):
        self.assertTrue(validate_request(draft()))
        r = request()
        r['cases'][0]['v_kn_per_m'] = None
        with self.assertRaises(ValueError):
            envelope(r)

    def test_signed_actions_use_independent_absolute_envelopes(self):
        self.assertEqual(envelope(request()), {'v': 80.4, 'm': 74.1, 'v_case': 'A', 'm_case': 'B'})

    def test_zero_valid_but_nonfinite_and_duplicate_cases_invalid(self):
        r = request()
        r['cases'][0]['v_kn_per_m'] = 0
        self.assertEqual(validate_request(r), [])
        r['cases'][1]['name'] = 'A'
        r['yield_strength_mpa'] = float('nan')
        r['stiffness_reduction'] = 1.2
        self.assertEqual(len(validate_request(r)), 3)
