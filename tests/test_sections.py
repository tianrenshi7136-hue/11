import unittest
from els.sections import properties, SECTIONS


class SectionTableTests(unittest.TestCase):
    def test_engineer_table_ei_and_units(self):
        p = properties('FSP IV', .5)
        self.assertEqual(p['wall_inertia_cm4_per_m'], 38600)
        self.assertEqual(p['ei_knm2_per_m'], 79130)
        self.assertEqual(p['reduced_ei_knm2_per_m'], 39565)
        self.assertEqual(properties('FSP VIL')['wall_mass_kg_per_m2'], 240)
        self.assertEqual(properties('FSP VL')['thickness_mm'], 24.3)

    def test_missing_invalid_section_or_bd(self):
        for section, bd in [('', .5), ('UNKNOWN', .5), ('FSP IV', 0), ('FSP IV', float('nan'))]:
            with self.assertRaises(ValueError):
                properties(section, bd)
        self.assertEqual(len(SECTIONS), 6)
