"""Continuous sheetpile checks using the engineer's existing Excel formulas."""
from contextlib import contextmanager
from datetime import datetime
import json
import hashlib
import math
from pathlib import Path
import shutil
import uuid

from .core import finite
from .templates import sha256

SECTIONS = ['FSP I', 'FSP II', 'FSPIII', 'FSP IV', 'FSP VL', 'FSP VIL']
STATUS_FORMULA = '=IF(AND(X50>T50,M50>I50),"OK","NOT OK")'
EXCEL_ERRORS = {'#REF!', '#DIV/0!', '#VALUE!', '#N/A', '#NAME?', '#NUM!', '#NULL!', '#SPILL!', '#CALC!', '#GETTING_DATA'}


def draft():
    return {'selection_mode': 'manual', 'section': '', 'yield_strength_mpa': None,
            'shear_reduction': None, 'bending_reduction': None, 'stiffness_reduction': None,
            'analysis_ei_knm2_per_m': None, 'cases': [], 'source': '', 'parameter_reference': ''}


def validate_request(request):
    if not isinstance(request, dict):
        return ['钢板桩输入须为对象']
    errors = []
    if request.get('selection_mode') not in ('manual', 'lightest'):
        errors.append('请选择手动型号或最轻合格截面')
    if request.get('selection_mode') == 'manual' and request.get('section') not in SECTIONS:
        errors.append('请选择模板中的钢板桩型号')
    for key in ('yield_strength_mpa', 'analysis_ei_knm2_per_m'):
        if not finite(request.get(key)) or request[key] <= 0:
            errors.append(f'{key}须填写正数')
    for key in ('shear_reduction', 'bending_reduction', 'stiffness_reduction'):
        if not finite(request.get(key)) or not 0 < request[key] <= 1:
            errors.append(f'{key}须在0到1之间且大于0')
    for key in ('source', 'parameter_reference'):
        if not isinstance(request.get(key), str) or not request[key].strip():
            errors.append(f'请填写{key}：分析来源或项目参数依据')
    cases = request.get('cases')
    if not isinstance(cases, list) or not cases:
        return errors + ['至少填写一个钢板桩分析工况']
    seen = set()
    for i, case in enumerate(cases, 1):
        if not isinstance(case, dict):
            errors.append(f'工况{i}格式无效')
            continue
        name = case.get('name')
        if not isinstance(name, str) or not name.strip() or name.strip() in seen:
            errors.append(f'工况{i}名称缺失或重复')
        else:
            seen.add(name.strip())
        for key in ('v_kn_per_m', 'm_knm_per_m'):
            if not finite(case.get(key)):
                errors.append(f'工况{i}的{key}须为有限数值；空值不能视为0')
    return errors


def envelope(request):
    errors = validate_request(request)
    if errors:
        raise ValueError('\n'.join(errors))
    cases = request['cases']
    v = max(cases, key=lambda c: abs(c['v_kn_per_m']))
    m = max(cases, key=lambda c: abs(c['m_knm_per_m']))
    return {'v': abs(v['v_kn_per_m']), 'm': abs(m['m_knm_per_m']),
            'v_case': v['name'], 'm_case': m['name']}


@contextmanager
def excel_session(path):
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    app = None
    wb = None
    try:
        app = win32com.client.DispatchEx('Excel.Application')
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.AutomationSecurity = 3
        app.AskToUpdateLinks = False
        app.Workbooks.Add()
        app.Calculation = -4135
        wb = app.Workbooks.Open(str(Path(path).resolve()), UpdateLinks=0, ReadOnly=False, AddToMru=False)
        yield app, wb
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        finally:
            try:
                if app is not None:
                    app.Quit()
            finally:
                pythoncom.CoUninitialize()


def check_template(wb):
    sheet = wb.Worksheets('FSPiv Summary')
    for address, expected in [('AB50', STATUS_FORMULA), ('I50', '=F50*1.4'), ('T50', '=Q50*1.4')]:
        if sheet.Range(address).Formula != expected:
            raise ValueError(f'模板{address}公式变化，须重新接入审查')
    for name in ('FSPiv Summary', 'Stiffness for FREW-FSPIV'):
        cell = wb.Worksheets(name).Range('J8')
        if cell.Validation.Type != 3 or cell.Validation.Formula1 != '=$AK$9:$AK$14':
            raise ValueError('Wall Property下拉映射变化')
        choices = [row[0] for row in wb.Worksheets(name).Range('AK9:AK14').Value2]
        if choices != SECTIONS:
            raise ValueError('截面库与已接入型号不一致')


def calculate_candidate(wb, section, request, actions):
    sheet = wb.Worksheets('FSPiv Summary')
    stiffness = wb.Worksheets('Stiffness for FREW-FSPIV')
    # Both sheets have independent dropdowns. Preserve their existing lookups.
    for ws in (sheet, stiffness):
        ws.Range('J8').Value2 = section
        ws.Range('J16').Value2 = request['yield_strength_mpa']
    # Project parameters explicitly replace historic project constants.
    for address, key in [('H40', 'yield_strength_mpa'), ('H44', 'yield_strength_mpa'),
                         ('S40', 'shear_reduction'), ('Q44', 'bending_reduction'),
                         ('R33', 'bending_reduction'), ('T33', 'stiffness_reduction')]:
        sheet.Range(address).Value2 = request[key]
    sheet.Range('F50').Value2 = actions['v']
    sheet.Range('Q50').Value2 = actions['m']
    # Use the same selected section and project reduction for the stiffness that
    # the summary reads, while retaining the original section lookup formulas.
    stiffness.Range('R35').Value2 = request['bending_reduction']
    stiffness.Range('T35').Value2 = request['stiffness_reduction']
    # Original V35 reads L29 (fixed >1 support-layer coefficient). The selected
    # project reduction must drive this output, rather than that historic row.
    stiffness.Range('V35').Formula = '=P35*T35'
    stiffness.Calculate()
    sheet.Calculate()
    result = {'section': section}
    for key, address in [('design_shear', 'I50'), ('design_moment', 'T50'),
                         ('shear_capacity', 'M50'), ('moment_capacity', 'X50'),
                         ('ei_knm2_per_m', 'V33'), ('weight_kg_per_m2', 'R21')]:
        cell = sheet.Range(address)
        if str(cell.Text) in EXCEL_ERRORS or not finite(cell.Value2):
            raise ValueError(f'{section} {address}公式报错或未返回数值：{cell.Text}')
        result[key] = float(cell.Value2)
    if min(result['shear_capacity'], result['moment_capacity'], result['ei_knm2_per_m'], result['weight_kg_per_m2']) <= 0:
        raise ValueError('承载力、刚度或重量必须为正数')
    result['status'] = sheet.Range('AB50').Value2
    expected = result['design_shear'] < result['shear_capacity'] and result['design_moment'] < result['moment_capacity']
    if result['status'] != ('OK' if expected else 'NOT OK'):
        raise ValueError('Excel通过标记与剪力／弯矩比较不一致')
    # Independently check the template's one-time action amplification.
    if not math.isclose(result['design_shear'], actions['v'] * 1.4, rel_tol=1e-10, abs_tol=1e-9) or not math.isclose(result['design_moment'], actions['m'] * 1.4, rel_tol=1e-10, abs_tol=1e-9):
        raise ValueError('模板内力放大与1.4约定不一致')
    result['shear_utilization'] = result['design_shear'] / result['shear_capacity']
    result['moment_utilization'] = result['design_moment'] / result['moment_capacity']
    width, height, thickness = [float(sheet.Range(c).Value2) for c in ('J12', 'J13', 'J14')]
    modulus = float(sheet.Range('R20').Value2)
    inertia = float(sheet.Range('R19').Value2)
    expected_shear = request['yield_strength_mpa'] * thickness * (height-thickness) / (width/1000) / math.sqrt(3) / 1000 * request['shear_reduction']
    expected_moment = request['yield_strength_mpa'] * modulus / 1000 * request['bending_reduction']
    expected_ei = 205 * inertia / 100 * request['stiffness_reduction']
    for key, expected_value in [('shear_capacity', expected_shear), ('moment_capacity', expected_moment), ('ei_knm2_per_m', expected_ei)]:
        if not math.isclose(result[key], expected_value, rel_tol=1e-9):
            raise ValueError(f'{key}与独立模板算式核对不一致')
    result['requires_reanalysis'] = not math.isclose(result['ei_knm2_per_m'], request['analysis_ei_knm2_per_m'], rel_tol=1e-6)
    return result


def generate(request, mapping_path, output_root, project_name):
    actions = envelope(request)
    if not isinstance(project_name, str) or not project_name.strip():
        raise ValueError('项目名称不能为空')
    mapping = json.loads(Path(mapping_path).read_text(encoding='utf-8'))
    template = Path(mapping['file'])
    before = sha256(template)
    if before != mapping['sha256']:
        raise ValueError('模板指纹变化，请重新检查及登记模板')
    folder = Path(output_root) / (datetime.now().strftime('%Y%m%d-%H%M%S') + '-' + uuid.uuid4().hex[:6])
    folder.mkdir(parents=True, exist_ok=False)
    target = folder / '钢板桩结构计算_待复核.xls'
    pdf = folder / '钢板桩结构计算_待复核.pdf'
    shutil.copy2(template, target)
    try:
        with excel_session(target) as (app, wb):
            check_template(wb)
            options = SECTIONS if request['selection_mode'] == 'lightest' else [request['section']]
            candidates = [calculate_candidate(wb, section, request, actions) for section in options]
            passed = [c for c in candidates if c['status'] == 'OK']
            if not passed:
                raise ValueError('无合格截面，未生成计算书或PDF；请调整候选范围或重新设计')
            selected = min(passed, key=lambda c: (c['weight_kg_per_m2'], SECTIONS.index(c['section'])))
            selected = calculate_candidate(wb, selected['section'], request, actions)
            sheet = wb.Worksheets('FSPiv Summary')
            for ws in wb.Worksheets:
                ws.Visible = -1 if ws.Name == 'FSPiv Summary' else 2
            sheet.Range('D3').Value2 = "'" + project_name
            sheet.Range('D4').Value2 = f"连续钢板桩墙 | V:{actions['v_case']} M:{actions['m_case']}"
            for address in ('R3', 'V3', 'AB3', 'AG3', 'V4', 'AB4'):
                sheet.Range(address).MergeArea.ClearContents()
            sheet.Range('B48').Formula = '=J8'
            sheet.Range('B50').Value2 = '模型内力绝对值包络'
            sheet.Range('J45').Value2 = 'kNm/m'
            sheet.Range('I48').Value2 = '1.4V (kN/m)'
            sheet.Range('T48').Value2 = '1.4M (kNm/m)'
            sheet.Range('Q48').Value2 = '模型弯矩 (kNm/m)'
            sheet.Range('M49').Value2 = '(kN/m)'
            sheet.Range('X49').Value2 = '(kNm/m)'
            stiffness = wb.Worksheets('Stiffness for FREW-FSPIV')
            for address, formula in [('H40', '=J16'), ('H44', '=J16'), ('R33', '=Q44')]:
                sheet.Range(address).Formula = formula
            for address, formula in [('J8', "='FSPiv Summary'!J8"), ('J16', "='FSPiv Summary'!J16"),
                                     ('R35', "='FSPiv Summary'!Q44"), ('T35', "='FSPiv Summary'!T33")]:
                stiffness.Range(address).Formula = formula
            # Exclude historic lagging cases and the section table from output.
            sheet.Rows('51:65536').Hidden = True
            sheet.Rows('24:29').Hidden = True
            sheet.Rows('34:34').Hidden = True
            for address in ('F50', 'Q50', 'I50', 'T50', 'M50', 'X50'):
                sheet.Range(address).NumberFormat = '0.00'
            notes = [
                '适用范围：仅连续钢板桩墙剪力及弯矩验算，待工程师复核；未包含整体稳定、节点或图纸。',
                '分析来源：' + request['source'],
                '项目材料及折减参数依据：' + request['parameter_reference'],
                f"模型输出未乘1.4，模板乘1.4一次；控制工况 V:{actions['v_case']} / M:{actions['m_case']}",
                f"分析模型EI={request['analysis_ei_knm2_per_m']:g}；选型后EI={selected['ei_knm2_per_m']:g} kN·m²/m。" + ('须外部重新分析。' if selected['requires_reanalysis'] else '与录入的分析模型刚度一致。'),
            ]
            for row, note in enumerate(notes, 51):
                cells = sheet.Range(f'A{row}:AH{row}')
                cells.UnMerge()
                cells.Clear()
                cells.Merge()
                cells.Value2 = note
                cells.WrapText = True
                cells.Font.Name = 'Microsoft YaHei'
                cells.Font.Size = 9
                cells.VerticalAlignment = -4108
                sheet.Rows(f'{row}:{row}').RowHeight = 30
            sheet.Rows('51:55').Hidden = False
            sheet.Range('A55:AH55').Borders(9).LineStyle = 1
            sheet.PageSetup.PrintArea = '$A$1:$AH$55'
            sheet.PageSetup.Zoom = False
            sheet.PageSetup.FitToPagesWide = 1
            sheet.PageSetup.FitToPagesTall = 1
            sheet.PageSetup.CenterHeader = '待工程师复核 / REVIEW DRAFT'
            sheet.PageSetup.CenterFooter = '刚度变化，须重新分析' if selected['requires_reanalysis'] else '仅连续钢板桩剪力及弯矩验算'
            app.Calculation = -4105
            stiffness.Calculate()
            sheet.Calculate()
            # Only inspect the printed, applicable formula area; historic
            # hidden cases may contain unrelated errors and are not endorsed.
            values = sheet.Range('A1:AH50').Value2
            formulas = sheet.Range('A1:AH50').Formula
            for row, (row_values, row_formulas) in enumerate(zip(values, formulas), 1):
                if row in range(24, 30) or row == 34:
                    continue
                for col, (value, formula) in enumerate(zip(row_values, row_formulas), 1):
                    if isinstance(formula, str) and formula.startswith('='):
                        cell = sheet.Cells(row, col)
                        if str(cell.Text) in EXCEL_ERRORS:
                            raise ValueError(f'输出公式错误：{cell.Address} {cell.Text}')
            wb.Save()
            sheet.ExportAsFixedFormat(0, str(pdf.resolve()))
        if sha256(template) != before:
            raise ValueError('原模板在生成期间变化，结果不能登记为有效')
        result = {'module': 'sheetpile_structure', 'review_status': 'REANALYSIS_REQUIRED' if selected['requires_reanalysis'] else 'ENGINEER_REVIEW_REQUIRED',
                  'project_name': project_name, 'request': request, 'envelope': actions,
                  'template_sha256': before, 'selected': selected, 'candidates': candidates,
                  'input_sha256': hashlib.sha256(json.dumps({'project_name': project_name, 'request': request}, ensure_ascii=False, sort_keys=True, allow_nan=False).encode('utf-8')).hexdigest(),
                  'workbook': str(target.resolve()), 'pdf': str(pdf.resolve()),
                  'scope': 'Only continuous sheetpile shear and bending checks; no global stability, connection, or CAD approval.',
                  'historical_sheets': 'Original unrelated worksheets retained very hidden for formula preservation; excluded from delivered PDF and assessment.'}
        result['controlled_formula_updates'] = {
            'Stiffness for FREW-FSPIV!V35': '=P35*T35 (project reduction replaces historic layer-table dependency)',
            'Stiffness for FREW-FSPIV!J8/J16/R35/T35': 'Linked to selected summary section/material/reductions',
            'FSPiv Summary!H40/H44/R33': 'Linked to project yield strength and bending reduction',
        }
        (folder / 'result.json').write_text(json.dumps(result, ensure_ascii=False, indent=2, allow_nan=False), encoding='utf-8')
        return result
    except Exception as exc:
        # Remove only newly-created output files; leave original and old results.
        for p in (target, pdf):
            if p.exists():
                p.unlink()
        (folder / 'failure.json').write_text(json.dumps({'error': str(exc), 'request': request}, ensure_ascii=False, indent=2), encoding='utf-8')
        raise
