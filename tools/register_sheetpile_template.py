"""Register the engineer-selected local template and verified input convention."""
import json
from pathlib import Path

root = Path(__file__).resolve().parent.parent
reports = json.loads((root / 'outputs/twd-xls-formulas.json').read_text(encoding='utf-8'))
r = next(r for r in reports if Path(r['path']).name == '4.7m-Section+Member+checking.xls')
sheet = next(s for s in r['sheets'] if s['name'] == 'FSPiv Summary')
formulas = {f['cell']: f['formula'] for f in sheet['formulas']}
expected = '=IF(AND(X50>T50,M50>I50),"OK","NOT OK")'
if formulas.get('AB50') != expected or formulas.get('I50') != '=F50*1.4' or formulas.get('T50') != '=Q50*1.4':
    raise RuntimeError('Template calculation or status formula changed; inspect again')
mapping = {
    'schema_version': 1, 'module': 'sheetpile_structure',
    'selected_by_engineer': True, 'selected_on': '2026-10-13',
    'baseline_validation_complete': False,
    'file': r['path'], 'sha256': r['sha256'],
    'sheet': 'FSPiv Summary', 'scope': 'continuous_sheetpile_wall_Type_C_row_50',
    'selection': {'cell': 'J8', 'method': 'existing_wall_property_dropdown_and_formula_dependencies'},
    'inputs': {'model_shear_kn_per_m': 'F50', 'model_moment_knm_per_m': 'Q50'},
    'load_convention': {
        'model_actions_factored': False, 'template_action_multiplier': 1.4,
        'apply_multiplier_in_software': False,
        'description': 'Steel sheetpile model V/M are unmultiplied; existing Excel applies 1.4 once. This is separate from the stability FoS target of 1.5.',
        'strut_waler_convention': 'Earlier factored inputs excluding member selfweight remain unchanged.'
    },
    'outputs': {'design_shear': 'I50', 'design_moment': 'T50', 'shear_capacity': 'M50', 'moment_capacity': 'X50', 'status': 'AB50'},
    'expected_formulas': {'I50': formulas['I50'], 'T50': formulas['T50'], 'AB50': formulas['AB50']},
    'project_parameter_policy': 'Yield strength, shear/bending/stiffness reduction must be explicitly supplied for each project. No historic defaults.',
    'excluded_rows': [53, 54, 55],
    'pending': ['Excel recalculation and changed-load tests', 'Dropdown choice enumeration and cross-sheet selection dependency test', 'Engineering confirmation of material/reduction applicability', 'CAD mapping and corresponding project analysis inputs'],
}
folder = root / 'templates'
folder.mkdir(exist_ok=True)
(folder / 'sheetpile-structure.local.json').write_text(json.dumps(mapping, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
text = '''# 钢板桩模板选定及输入口径

日期：2026-10-13。

- 工程师选定：TWD-TD004/design/4.7m-Section+Member+checking.xls。
- 首个接入范围：FSPiv Summary第50行连续钢板桩墙Type C；其他lagging及FSP III组合另行处理。
- 保留Wall Property原下拉菜单及其已有参数更新公式。下拉菜单是输入入口，VLOOKUP是参数联动公式，软件无需再建立第二套属性计算。
- 钢板桩剪力、弯矩输入为模型未乘1.4的输出，单位分别为kN/m、kNm/m；模板乘1.4一次，软件不提前乘。
- 工程师确认材料强度和剪力、弯矩、刚度折减系数按项目设置，不采用原模板数值作为默认值。
- 此约定更新此前对钢板桩输入已分项的假设；支撑、围檩仍沿用已分项且未含构件自重的已确认口径。
- 1.4登记为模板内力放大系数，与整体稳定目标FoS 1.5分开，不自动解释为全部规范安全要求。
- 本次重新读取确认AB50已改为IF(AND(X50>T50,M50>I50),"OK","NOT OK")，剪力与弯矩同时满足才通过；保留原表严格大于的边界规则。
- AB53、AB54、AB55当前仍只判弯矩，不纳入首个自动生成模块。第55行标题写FSP III，但当前承载力仍沿用FSP IV数值，需单独核查后再扩展。
- 已生成本地输入输出映射及模板SHA256，工程模板不上传公开仓库。
- 尚未完成Excel重算、下拉型号切换、超限及缺输入测试，模板选定不等于正式成果生成验证完成。
'''
(root / 'outputs/钢板桩模板选定记录.md').write_text(text, encoding='utf-8')
print('Local mapping and selection record saved')
