"""Engineer-supplied FSP table, transcribed from the provided image."""
FIELDS = ('width_mm', 'height_mm', 'thickness_mm', 'pile_area_cm2', 'pile_inertia_cm4',
          'pile_modulus_cm3', 'pile_mass_kg_per_m', 'wall_area_cm2_per_m',
          'wall_inertia_cm4_per_m', 'wall_modulus_cm3_per_m', 'wall_mass_kg_per_m2')
ROWS = {
    'FSP I': (400, 85, 8, 45.21, 598, 88, 35.5, 113, 4500, 529, 88.8),
    'FSP II': (400, 100, 10.5, 61.18, 1240, 152, 48, 153, 8740, 874, 120),
    'FSPIII': (400, 125, 13, 76.42, 2220, 223, 60, 191.5, 16800, 1340, 150),
    'FSP IV': (400, 170, 15.5, 96.99, 4670, 362, 76.1, 242.5, 38600, 2270, 190),
    'FSP VL': (500, 200, 24.3, 133.8, 7960, 520, 105, 267.6, 63000, 3150, 210),
    'FSP VIL': (500, 225, 27.6, 153, 11400, 680, 120, 306, 86000, 3820, 240),
}
SECTIONS = list(ROWS)
E_KN_PER_MM2 = 205.0
YIELD_MPA = 275.0
TABLE_VERSION = 'engineer-image-20261013-v1'


def properties(section, bd=None):
    if section not in ROWS:
        raise ValueError('请选择截面库中的钢板桩型号')
    result = dict(zip(FIELDS, ROWS[section]))
    result.update(section=section, yield_strength_mpa=YIELD_MPA,
                  elastic_modulus_kn_per_mm2=E_KN_PER_MM2,
                  ei_knm2_per_m=E_KN_PER_MM2 * result['wall_inertia_cm4_per_m'] / 100)
    if bd is not None:
        if isinstance(bd, bool) or not isinstance(bd, (int, float)) or not 0 < bd <= 1:
            raise ValueError('EI折减系数Bd须大于0且不超过1')
        result['reduced_ei_knm2_per_m'] = result['ei_knm2_per_m'] * bd
    return result
