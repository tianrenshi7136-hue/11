"""Integration checks against the selected local template, not new design data."""
import copy
import json
from pathlib import Path
import shutil
from els.sheetpile import SECTIONS, calculate_candidate, check_template, envelope, excel_session, generate

root = Path(__file__).resolve().parent.parent
mapping = root / 'templates/sheetpile-structure.local.json'
config = json.loads(mapping.read_text(encoding='utf-8'))
request = {'selection_mode': 'manual', 'section': 'FSP IV', 'yield_strength_mpa': 275,
           'shear_reduction': .7, 'bending_reduction': .7, 'stiffness_reduction': .5,
           'analysis_ei_knm2_per_m': 39565, 'source': 'TEMPLATE BASELINE TEST ONLY',
           'parameter_reference': 'HISTORIC TEMPLATE VALUES FOR SOFTWARE TEST ONLY',
           'cases': [{'name': 'BASELINE TEST', 'v_kn_per_m': 80.4, 'm_knm_per_m': 74.1}]}
scratch = root / 'outputs/sheetpile-integration-copy.xls'
shutil.copy2(config['file'], scratch)
results = []
with excel_session(scratch) as (app, wb):
    check_template(wb)
    for section in SECTIONS:
        result = calculate_candidate(wb, section, request, envelope(request))
        results.append(result)
    baseline = next(r for r in results if r['section'] == 'FSP IV')
    assert abs(baseline['shear_capacity'] - 665.3808493745183) < 1e-6
    assert abs(baseline['moment_capacity'] - 436.975) < 1e-6
    assert not baseline['requires_reanalysis']
    # Independently exceed only one action at a time: both must fail.
    for v, m in [(baseline['shear_capacity']/1.4*1.01, 0), (0, baseline['moment_capacity']/1.4*1.01)]:
        assert calculate_candidate(wb, 'FSP IV', request, {'v': v, 'm': m})['status'] == 'NOT OK'
    assert calculate_candidate(wb, 'FSP IV', request, {'v': 0, 'm': 0})['status'] == 'OK'
    changed = copy.deepcopy(request)
    changed.update(yield_strength_mpa=355, shear_reduction=.6, bending_reduction=.8, stiffness_reduction=.35)
    result = calculate_candidate(wb, 'FSP IV', changed, envelope(changed))
    assert result['requires_reanalysis']
    assert abs(result['ei_knm2_per_m'] - 79130 * .35) < 1e-6
    results.append(result)
scratch.unlink()
output = generate(request, mapping, root / 'outputs/sheetpile-validation', 'Software Validation Sample - Not a Project Design')
failed = copy.deepcopy(request)
failed['cases'][0].update(v_kn_per_m=1e8, m_knm_per_m=1e8)
failure_root = root / 'outputs/sheetpile-failure-validation'
try:
    generate(failed, mapping, failure_root, 'NO QUALIFIED SECTION TEST')
except ValueError as exc:
    assert '未通过验算' in str(exc)
else:
    raise AssertionError('Failed design unexpectedly generated')
assert not list(failure_root.rglob('*.xls'))
assert not list(failure_root.rglob('*.pdf'))
(root / 'outputs/sheetpile-integration-results.json').write_text(json.dumps({'checks': results, 'generated': output, 'status': 'PASS'}, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps({'status': 'PASS', 'workbook': output['workbook'], 'pdf': output['pdf']}, ensure_ascii=False))
