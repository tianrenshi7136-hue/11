"""Continuous sheetpile checks using the engineer's existing Excel formulas."""
from contextlib import contextmanager
from datetime import datetime
import json
import hashlib
import math
from pathlib import Path
import shutil
import uuid
import tempfile
import re

from .core import finite
from .templates import sha256
from .sections import SECTIONS, ROWS, TABLE_VERSION, properties

STATUS_FORMULA = '=IF(AND(X50>T50,M50>I50),"OK","NOT OK")'
EXCEL_ERRORS = {'#REF!', '#DIV/0!', '#VALUE!', '#N/A', '#NAME?', '#NUM!', '#NULL!', '#SPILL!', '#CALC!', '#GETTING_DATA'}


def excel_call(action):
    """Excel can briefly reject automation calls while finishing a save/export."""
    import pythoncom,pywintypes,time
    for attempt in range(24):
        try:return action()
        except pywintypes.com_error as exc:
            if exc.hresult not in (-2147418111,-2147417846) or attempt==23:raise
            pythoncom.PumpWaitingMessages();time.sleep(.25)


def save_and_export(wb,pdf):
    excel_call(wb.__dict__.get('_els_save',lambda:wb.Save()))
    export=wb.__dict__.get('_els_export')
    excel_call(lambda:export(0,str(pdf)) if export else wb.ExportAsFixedFormat(0,str(pdf)))


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
    keys = ('yield_strength_mpa',) if request.get('purpose') == 'preselection' else ('yield_strength_mpa', 'analysis_ei_knm2_per_m')
    for key in keys:
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


def rejection_message(candidates, mode):
    heading = (f"当前指定型号 {candidates[0]['section']} 未通过验算（只检查该型号，未自动尝试其他型号）。"
               if mode == 'manual' and candidates else '候选库中无合格截面。')
    lines = [heading, '未生成计算书或PDF。以下数值为1.4倍模型内力／承载力：']
    for c in candidates:
        lines.extend([c['section'],
            f"  剪力：{c['design_shear']:.2f} / {c['shear_capacity']:.2f} kN/m；利用率{c['shear_utilization']:.1%}，{'通过' if c['shear_utilization'] < 1 else '未通过'}",
            f"  弯矩：{c['design_moment']:.2f} / {c['moment_capacity']:.2f} kNm/m；利用率{c['moment_utilization']:.1%}，{'通过' if c['moment_utilization'] < 1 else '未通过'}"])
    return '\n'.join(lines)


@contextmanager
def excel_session(path):
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    app = None
    wb = None
    close_book = None
    quit_app = None
    try:
        app = win32com.client.DispatchEx('Excel.Application')
        quit_app = app.Quit
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.AutomationSecurity = 3
        app.AskToUpdateLinks = False
        app.Workbooks.Add()
        app.Calculation = -4135
        try:
            wb = app.Workbooks.Open(str(Path(path).resolve()), UpdateLinks=0, ReadOnly=False, AddToMru=False)
            close_book = wb.Close
            # Bind these while Excel is idle. Late-bound type lookup itself can
            # fail during the short busy interval after Save/Export.
            wb.__dict__['_els_save']=wb.Save
            wb.__dict__['_els_export']=wb.ExportAsFixedFormat
        except Exception as exc:
            raise ValueError(f'Excel 无法打开工作副本：{Path(path).resolve()}。这是文件读取失败，尚未执行验算。\n{exc}') from exc
        yield app, wb
    finally:
        try:
            if close_book is not None:
                excel_call(lambda:close_book(SaveChanges=False))
        finally:
            try:
                if quit_app is not None:
                    excel_call(quit_app)
            finally:
                pythoncom.CoUninitialize()


def check_template(wb):
    from .wall_compat import prepare_wall
    prepare_wall(wb)
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
        # Update only the generation copy; existing VLOOKUPs remain authoritative
        # calculations but now read the same engineer-supplied table as the UI.
        ws.Range('AL9:AV14').Value2 = tuple(ROWS[name] for name in SECTIONS)
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
    sheet.Calculate()
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
    reference = request.get('analysis_ei_knm2_per_m')
    result['requires_reanalysis'] = finite(reference) and not math.isclose(result['ei_knm2_per_m'], reference, rel_tol=1e-6)
    return result


def preselect(request, mapping_path, output_root):
    if request.get('purpose') != 'preselection' or request.get('selection_mode') != 'lightest':
        raise ValueError('截面初选须使用独立初选输入')
    actions = envelope(request)
    from .wall_compat import wall_reference
    template=wall_reference(mapping_path)
    revision=sha256(template)
    with tempfile.TemporaryDirectory() as folder:
        copy = Path(folder) / ('preselection'+template.suffix)
        shutil.copy2(template, copy)
        with excel_session(copy) as (app, wb):
            check_template(wb)
            candidates = [calculate_candidate(wb, section, request, actions) for section in SECTIONS]
    for c in candidates:
        c.pop('requires_reanalysis', None)
    passed = [c for c in candidates if c['status'] == 'OK']
    result = {'purpose': 'preselection', 'request': request, 'envelope': actions, 'candidates': candidates,
              'selected': min(passed, key=lambda c: c['weight_kg_per_m2']) if passed else None,
              'status': 'PRELIMINARY_SELECTION_ONLY', 'template_sha256': revision}
    output = Path(output_root) / (datetime.now().strftime('%Y%m%d-%H%M%S') + '-' + uuid.uuid4().hex[:6])
    output.mkdir(parents=True)
    (output / 'preselection.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    return result


def generate(request, mapping_path, output_root, project_name, subject_name='Continuous Sheet Pile Wall Structural Check'):
    if request.get('selection_mode') != 'manual' or request.get('purpose') == 'preselection':
        raise ValueError('计算书生成只验证指定型号；请在独立截面初选页面比较候选')
    actions = envelope(request)
    support_package=request.get('support_package')
    if support_package:
        from .supports import validate_package, section_library
        errors=validate_package(support_package,section_library(aa=True))
        if errors:
            raise ValueError('\n'.join(errors))
    if not isinstance(project_name, str) or not project_name.strip():
        raise ValueError('项目名称不能为空')
    if not isinstance(subject_name, str) or not subject_name.strip():
        raise ValueError('Subject Name不能为空')
    if re.search(r'[\u3400-\u9fff]', project_name + subject_name):
        raise ValueError('Project Name和Subject Name请使用英文，不得含中文字符')
    from .reference_store import detach_external_links
    from .wall_compat import wall_reference
    template=wall_reference(mapping_path)
    before = sha256(template)
    folder = Path(output_root) / (datetime.now().strftime('%Y%m%d-%H%M%S') + '-' + uuid.uuid4().hex[:6])
    folder.mkdir(parents=True, exist_ok=False)
    output_name='ELS_Structural_Check_Review' if support_package else 'Sheetpile_Structural_Check_Review'
    target = folder / (output_name+template.suffix)
    pdf = folder / (output_name+'.pdf')
    shutil.copy2(template, target)
    candidates = []
    try:
        with excel_session(target) as (app, wb):
            if support_package and template.suffix.lower()=='.xlsx':
                from .connections import prepare_aa_workbook
                prepare_aa_workbook(wb)
            check_template(wb)
            options = SECTIONS if request['selection_mode'] == 'lightest' else [request['section']]
            candidates = [calculate_candidate(wb, section, request, actions) for section in options]
            passed = [c for c in candidates if c['status'] == 'OK']
            selected = candidates[0]
            selected = calculate_candidate(wb, selected['section'], request, actions)
            sheet = wb.Worksheets('FSPiv Summary')
            for ws in wb.Worksheets:
                ws.Visible = -1 if ws.Name == 'FSPiv Summary' else 2
            sheet.Range('D3').Value2 = "'" + project_name
            sheet.Range('D4').Value2 = "'" + subject_name
            sheet.Range('Q4').NumberFormat='[$-409]dd-mmm-yy'
            sheet.Range('Q4').ShrinkToFit=True
            sheet.Range('Q4').Font.Size=8
            sheet.Range('H47').Value2='Actions = 1.4 x unfactored analysis'
            for address in ('R3', 'V3', 'AB3', 'AG3', 'V4', 'AB4'):
                sheet.Range(address).MergeArea.ClearContents()
            sheet.Range('B48').Formula = '=J8'
            sheet.Range('B50').Value2 = 'Model Action Envelope'
            sheet.Range('J45').Value2 = 'kNm/m'
            sheet.Range('I48').Value2 = '1.4V (kN/m)'
            sheet.Range('T48').Value2 = '1.4M (kNm/m)'
            sheet.Range('Q48').Value2 = 'Model Moment (kNm/m)'
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
            sheet.Range('A51:AH55').UnMerge()
            sheet.Range('A51:AH55').Clear()
            support_result=None
            output_sheets=[sheet.Name]
            if support_package:
                from .supports import populate
                extra,support_result=populate(wb,support_package,actions,project_name,subject_name)
                output_sheets.extend(extra)
            for sheet_name in reversed(output_sheets):
                if wb.Worksheets(1).Name != sheet_name:
                    wb.Worksheets(sheet_name).Move(wb.Worksheets(1))
            # English-only workbook metadata, including hidden historic sheets.
            # CELL(filename) otherwise introduces Chinese directory/file names.
            for ws in wb.Worksheets:
                used = ws.UsedRange
                formulas = used.Formula
                row_start, col_start = used.Row, used.Column
                if isinstance(formulas, tuple):
                    for ri, row_values in enumerate(formulas):
                        for ci, formula in enumerate(row_values):
                            if isinstance(formula, str) and 'CELL("FILENAME"' in formula.upper():
                                ws.Cells(row_start+ri, col_start+ci).Value2 = f'OUTPUT/[{target.name}]{ws.Name}'
                ws.PageSetup.LeftHeader = ''
                ws.PageSetup.RightHeader = ''
                ws.PageSetup.CenterHeader = 'REVIEW DRAFT - NOT PASS' if selected['status']!='OK' or (support_result and support_result.get('status')=='NOT PASS') else 'REVIEW DRAFT'
                ws.PageSetup.LeftFooter = ''
                ws.PageSetup.RightFooter = ''
                ws.PageSetup.CenterFooter = ''
            # Replace the bilingual header picture; retain the FSP section sketch.
            for index in range(sheet.Shapes.Count, 0, -1):
                shape = sheet.Shapes.Item(index)
                if shape.Type in (11, 13) and shape.TopLeftCell.Row <= 2:
                    shape.Delete()
            header = sheet.Range('A2:AH2')
            header.UnMerge()
            header.Merge()
            header.Value2 = 'STRUCTURAL CALCULATIONS'
            header.Font.Name = 'Arial'
            header.Font.Bold = True
            header.HorizontalAlignment = -4108
            from .printing import setup_print
            setup_print(sheet,50,app)
            app.Calculation = -4105
            stiffness.Calculate()
            sheet.Calculate()
            for sheet_name in output_sheets[1:]:
                checked=wb.Worksheets(sheet_name)
                checked.Calculate()
                area=checked.Range(checked.PageSetup.PrintArea)
                try:
                    errors=area.SpecialCells(-4123,16)
                except Exception:
                    errors=None
                if errors is not None:
                    raise ValueError(f'{sheet_name}打印区域存在公式错误：{errors.Address}')
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
            detach_external_links(wb,output_sheets)
            from .printing import format_check_results
            for name in output_sheets:format_check_results(wb.Worksheets(name))
            save_and_export(wb,pdf.resolve())
        result = {'module': 'sheetpile_structure', 'review_status': 'REANALYSIS_REQUIRED' if selected['requires_reanalysis'] else 'ENGINEER_REVIEW_REQUIRED',
                  'section_table_version': TABLE_VERSION,
                  'project_name': project_name, 'subject_name': subject_name, 'request': request, 'envelope': actions,
                  'output_sheets':output_sheets, 'support_package_results':support_result,
                  'template_sha256': before, 'selected': selected, 'candidates': candidates,
                  'input_sha256': hashlib.sha256(json.dumps({'project_name': project_name, 'subject_name': subject_name, 'request': request}, ensure_ascii=False, sort_keys=True, allow_nan=False).encode('utf-8')).hexdigest(),
                  'workbook': str(target.resolve()), 'pdf': str(pdf.resolve()),
                  'scope': ('Continuous sheetpile, template steel members, direct concentric welds and engineer-defined blinding checks; global stability and CAD pending.' if support_result else 'Only continuous sheetpile shear and bending checks; global stability, connections and CAD pending.'),
                  'historical_sheets': 'Original unrelated worksheets retained very hidden for formula preservation; excluded from delivered PDF and assessment.'}
        result['check_status']='NOT PASS' if selected['status']!='OK' or (support_result and support_result.get('status')=='NOT PASS') else 'PASS'
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
        (folder / 'failure.json').write_text(json.dumps({'error': str(exc), 'request': request, 'envelope': actions, 'candidates': candidates}, ensure_ascii=False, indent=2), encoding='utf-8')
        raise
