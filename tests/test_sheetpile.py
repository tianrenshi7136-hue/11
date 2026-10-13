import unittest
from els.sheetpile import draft, envelope, validate_request


def request():
    r = draft()
    r.update(section='FSP IV', yield_strength_mpa=275, shear_reduction=.7,
             bending_reduction=.7, stiffness_reduction=.5, analysis_ei_knm2_per_m=39565,
             source='TEST FREW', parameter_reference='TEST PARAMETERS',
             cases=[{'name': 'A', 'v_kn_per_m': -80.4, 'm_knm_per_m': 20},
                    {'name': 'B', 'v_kn_per_m': 5, 'm_knm_per_m': -74.1}])
    return r


class SheetpileTests(unittest.TestCase):
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
