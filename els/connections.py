"""Controlled AA-template connection blocks on one calculation worksheet."""
from contextlib import contextmanager
import json
import math
from pathlib import Path
import shutil
import tempfile

from .core import finite
from .templates import sha256
from .support_layout import copy_rows

AA_CONNECTIONS = 'aa-connections-v1'
BRACKET_FIELDS = ('bracket_id', 'case', 'waler_id', 'waler_case', 'section', 'grade',
                  'length_m', 'tributary_length_m', 'load_model', 'eccentricity_m',
                  'leg_mm', 'strength_mpa', 'effective_length_mm')


def validate_brackets(package, library=None):
    errors, seen = [], set()
    brackets = package.get('brackets', [])
    if not isinstance(brackets, list):
        return ['托架输入须为列表']
    members = {(m.get('member_id'), m.get('case')): m for m in package.get('members', []) if isinstance(m, dict)}
    import re
    for i, b in enumerate(brackets, 1):
        if not isinstance(b, dict):
            errors.append(f'托架{i}格式无效'); continue
        key = (b.get('bracket_id'), b.get('case'))
        if any(not isinstance(x, str) or not re.fullmatch(r'[A-Za-z0-9_-]{1,12}', x) for x in key):
            errors.append(f'托架{i}编号及工况须为1～12位英文字母／数字／短横线')
        elif key in seen:
            errors.append(f'托架{i}编号与工况重复')
        else:
            seen.add(key)
        waler = members.get((b.get('waler_id'), b.get('waler_case')))
        if not waler or waler.get('kind') != 'Waler':
            errors.append(f'托架{i}须关联有效围檩及工况')
        if not b.get('section') or (library is not None and b['section'] not in library):
            errors.append(f'托架{i}截面不在模板库')
        if b.get('grade') not in ('S235', 'S275', 'S355', 'S460'):
            errors.append(f'托架{i}钢材等级无效')
        for field in ('length_m', 'tributary_length_m', 'leg_mm', 'strength_mpa', 'effective_length_mm'):
            if not finite(b.get(field)) or b[field] <= 0:
                errors.append(f'托架{i} {field}须填写正数')
        if b.get('load_model') != 'point':
            errors.append(f'托架{i}须采用已确认的围檩重心集中力模型')
    return errors


def bracket_actions(bracket, waler_area_cm2, bracket_area_cm2):
    """Gravity loads only. Areas cm2, lengths m, forces kN, moments kNm."""
    length = bracket['length_m']
    waler_force = 1.4 * 77 * waler_area_cm2 / 10000 * bracket['tributary_length_m']
    selfweight = 1.4 * 77 * bracket_area_cm2 / 10000
    eccentricity = length / 2
    return dict(waler_force_kn=waler_force, selfweight_kn_per_m=selfweight,
                shear_kn=waler_force+selfweight*length,
                moment_knm=waler_force*eccentricity+selfweight*length**2/2)


@contextmanager
def aa_source(app):
    """Inspect a source snapshot; the approved engineer file is never opened writable."""
    from .supports import ROOT
    from .reference_store import reference_file
    source=reference_file(ROOT/'templates/supports-structure.local.json')
    expected=sha256(source)
    with controlled_source(app, source, expected) as opened:
        yield opened


@contextmanager
def controlled_source(app, source, expected):
    # A local snapshot also handles Excel/OneDrive readers that deny Python's
    # default sharing mode. PowerShell Copy-Item uses Windows file sharing.
    with tempfile.TemporaryDirectory(prefix='els-aa-') as folder:
        snapshot = Path(folder)/'aa-source.xlsx'
        copy_source(source, snapshot)
        book = app.Workbooks.Open(str(snapshot), UpdateLinks=0, ReadOnly=True, AddToMru=False)
        try:
            yield book, expected
        finally:
            book.Close(SaveChanges=False)


def copy_source(source, destination):
    try:
        shutil.copyfile(source, destination)
    except PermissionError:
        import os
        import subprocess
        env = dict(os.environ, ELS_TEMPLATE_SOURCE=str(Path(source).resolve()),
                   ELS_TEMPLATE_DEST=str(Path(destination).resolve()))
        subprocess.run(['powershell.exe', '-NoProfile', '-NonInteractive', '-Command',
                        'Copy-Item -LiteralPath $env:ELS_TEMPLATE_SOURCE -Destination $env:ELS_TEMPLATE_DEST -ErrorAction Stop'],
                       env=env, check=True, capture_output=True,
                       creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))


def prepare_aa_workbook(wb):
    """Adapt AA-native sheet names and summary columns to the common writer.

    No legacy XLS workbook is opened or copied. The member calculations and
    section library remain those in the engineer-selected AA workbook.
    """
    from .supports import report_row
    for old,new in [('CS1 ','CCS1 '),('W1','CW1'),('D-Strut&Wailing Summary','Strut&Wailing Summary')]:
        wb.Worksheets(old).Name=new
    summary=wb.Worksheets('Strut&Wailing Summary')
    summary.Range('A7:AH100').UnMerge();summary.Range('A7:AH100').Clear()
    strut_spans=[('B','C'),('D','H'),('I','K'),('L','N'),('O','Q'),('R','T'),('U','W'),('X','Z'),('AA','AB'),('AC','AE'),('AF','AH')]
    action_spans=[('B','C'),('D','G'),('H','J'),('K','L'),('M','N'),('O','P'),('Q','Q'),('R','T'),('U','V'),('W','W'),('X','Z'),('AA','AB'),('AC','AC'),('AD','AF'),('AG','AH')]
    summary.Range('B9').Value2='I. Strut Design'
    report_row(summary,10,['Mark','Section','Spacing (m)','Angle (deg)','Normal F (kN)','Adopted F (kN)','Design 1.4F (kN)','Axial N (kN)','','Capacity Pc (kN)','Check'],strut_spans,True)
    report_row(summary,12,['']*len(strut_spans),strut_spans)
    summary.Range('B16').Value2='II. Waling Design'
    report_row(summary,17,['Mark','Section','q (kN/m)','Load width (m)','Span (m)','Design N (kN)','','Capacity Pc (kN)','Design V (kN)','','Capacity Vc (kN)','Design M (kNm)','','Capacity Mc (kNm)','Check'],action_spans,True)
    report_row(summary,19,['']*len(action_spans),action_spans)
    for r in (9,16):summary.Rows(r).RowHeight=20;summary.Range(f'B{r}:AH{r}').Font.Bold=True
    for r in (11,18):summary.Rows(r).RowHeight=8
    wb.Names('Table').RefersTo="='Section Table UC'!$A$5:$X$128"


def guarded_status(capacity, demand):
    return (f'=IF(COUNT({capacity},{demand})<2,"CHECK INPUTS",'
            f'IF(AND({capacity}>0,{demand}>=0,{capacity}>={demand}),"OK","NOT OK"))')


def build_aa_connections(wb, sheet, package, member_sheets, numeric):
    """Use the requested workbook's blocks; replace all original project links."""
    sheet.Range('A7:AH500').UnMerge(); sheet.Range('A7:AH500').Clear()
    results, bracket_results, row = [], [], 7
    with aa_source(wb.Application) as (source_book, fingerprint):
        source = source_book.Worksheets('Connection check')
        # Match the supplied page: waling / sheet pile, brackets, strut / waling.
        from .critical_connections import governing_welds,attach_live_selector
        ordered = sorted(governing_welds(package,member_sheets), key=lambda w: w['kind'] != 'Waler-Sheetpile')
        def emit_weld(w):
            nonlocal row
            member, design = member_sheets[(w['member_id'], w['case'])]
            at = lambda c, offset: f'{c}{row+offset}'
            ref = lambda c: f"'{member.Name}'!{c}"
            label = w['member_id']+' / '+w['case']
            if w['kind'] == 'Waler-Sheetpile':
                copy_rows(source, sheet, 18, 23, row)
                sheet.Range(at('B',0)).Value2 = f'Connection Check between Waling and Sheet Piles ({label})'
                for c,offset,value in [('H',1,'kN from waling'),('Q',1,w['distribution_length_m']),('R',1,'m'),('S',1,'long'),('B',3,w['leg_mm']),('O',4,w['strength_mpa']),('U',4,w['effective_length_mm']/2),('Z',4,2)]:
                    sheet.Range(at(c,offset)).Value2=value
                formulas={at('F',1):'='+ref('J12'), at('F',2):f'=IF(COUNT({at("F",1)},{at("Q",1)})<2,"INPUTS REQUIRED",{at("F",1)}/{at("Q",1)})',
                          at('J',4):'='+at('B',3), at('G',5):f'=0.7*{at("J",4)}*{at("O",4)}*{at("U",4)}*{at("Z",4)}/1000', at('K',5):'='+at('F',2)}
                cap, force, status, count = at('G',5),at('K',5),at('AE',5),6
                length=w['effective_length_mm'];demand=design['design_actions']['v']/w['distribution_length_m'];units='kN/m'
                sheet.Range(at('H',2)).Value2='kN/m';sheet.Range(at('I',5)).Value2='kN/m';sheet.Range(at('M',5)).Value2='kN/m'
                sheet.Range(at('I',5)).ClearContents()
                sheet.Range(f'{at("M",5)}:{at("Q",5)}').Merge()
            else:
                copy_rows(source, sheet, 7, 16, row)
                sheet.Range(at('B',0)).Value2=f'Connection: {"Corner Strut" if design["member"]["angle_deg"]>0 else "Strut"} - Waling ({label})'
                sheet.Range(at('B',1)).Value2='Design axial force';sheet.Range(at('P',1)).Value2='kN'
                sheet.Range(at('B',2)).Value2='Connection design force'
                sheet.Rows(row+1).Hidden=True
                sheet.Range(at('B',2)).Value2='Shear Load From Corner Strut' if design['member']['angle_deg']>0 else 'Load From Strut'
                sheet.Range(f'{at("P",2)}:{at("Y",2)}').UnMerge();sheet.Range(f'{at("P",2)}:{at("Y",2)}').Clear();sheet.Range(at('P',2)).Value2='kN'
                sheet.Range(at('D',3)).Value2=w['leg_mm'];sheet.Range(at('U',8)).Value2=w['strength_mpa']
                formulas={at('M',1):'='+ref('J11'),at('M',2):'='+at('M',1),at('J',4):'='+ref('R17'),at('J',5):'='+ref('R17')+'-'+ref('R18'),at('J',6):'='+ref('R20'),at('J',7):'='+at('D',3),at('V',7):f'={at("J",7)}*0.7'}
                for offset in (4,5,6):
                    sheet.Range(at('O',offset)).Value2=2;formulas[at('V',offset)]=f'={at("J",offset)}*{at("O",offset)}'
                for c,offset in [('K',4),('N',5),('Q',6)]:formulas[at(c,8)]='='+at('V',offset)
                formulas[at('X',8)]='='+at('V',7)
                formulas[at('J',9)]=f'=({at("K",8)}+{at("N",8)}+{at("Q",8)})*{at("U",8)}*{at("X",8)}/1000'
                formulas[at('O',9)]='='+at('M',2)
                cap,force,status,count=at('J',9),at('O',9),at('AE',9),10
                length=2*(numeric(member,'R17')+numeric(member,'R17')-numeric(member,'R18')+numeric(member,'R20'))
                demand=design['design_actions']['n'];units='kN'
            formulas[status]=guarded_status(cap,force)
            comparison=at('J',5) if w['kind']=='Waler-Sheetpile' else at('N',9)
            formulas[comparison]=f'=IF(COUNT({cap},{force})<2,"?",IF({cap}>={force},">=","<"))'
            for cell, formula in formulas.items():sheet.Range(cell).Formula=formula
            for cell in (cap,force):sheet.Range(cell).NumberFormat='0.00';sheet.Range(cell).ShrinkToFit=True
            if w['kind']=='Strut-Waler':
                for offset in (1,2):sheet.Range(at('M',offset)).NumberFormat='0.00'
            sheet.Calculate();capacity=numeric(sheet,cap)
            if not math.isclose(capacity,.7*w['leg_mm']*w['strength_mpa']*length/1000,rel_tol=1e-9) or not math.isclose(numeric(sheet,force),demand,rel_tol=1e-9):
                raise ValueError('AA连接模板与独立焊缝算式不一致')
            attach_live_selector(wb,sheet,w,member_sheets,row)
            results.append(dict(input=w,demand=demand,capacity=capacity,units=units,effective_length_mm=length,status='PASS' if capacity>=demand else 'FAIL',capacity_cell=cap,force_cell=force,status_cell=status))
            row+=count+1
        for w in ordered:
            if w['kind']=='Waler-Sheetpile':emit_weld(w)
        if package.get('brackets'):
            row,bracket_results=build_brackets(wb,sheet,source_book,package,member_sheets,row,numeric)
        for w in ordered:
            if w['kind']=='Strut-Waler':emit_weld(w)
    return results,bracket_results,row-2,fingerprint


def build_brackets(wb, sheet, source_book, package, member_sheets, row, numeric):
    results=[]
    source=source_book.Worksheets('Connection check')
    table="'Section Table UC'!$A$5:$X$128"
    for b in package['brackets']:
        waler,_=member_sheets[(b['waler_id'],b['waler_case'])]
        original=wb.Worksheets('CW1');visibility=original.Visible;original.Visible=-1
        try:
            original.Copy(wb.Worksheets('FSPiv Summary'));calc=wb.Application.ActiveSheet
        finally:original.Visible=visibility
        calc.Name=f"B-{b['bracket_id']}-{b['case']}"
        calc.Range('A1:CI65').UnMerge();calc.Range('A1:CI65').Clear()
        source_book.Worksheets('Bracket').Range('A1:CI65').Copy(calc.Range('A1'))
        # Read formula text from the source, so cross-workbook Copy cannot turn
        # the original Table name into an external workbook reference.
        values=source_book.Worksheets('Bracket').Range('A1:CI65').Formula
        for ri,values_row in enumerate(values,1):
            for ci,f in enumerate(values_row,1):
                if isinstance(f,str) and f.startswith('='):
                    f=f.replace(',Table,',','+table+',').replace('INDEX(Table,','INDEX('+table+',')
                    calc.Cells(ri,ci).Formula=f
        calc.Range('P11').Value2=b['section'];calc.Range('Y11').Value2=b['grade']
        calc.Range('J16').Value2=b['length_m'];calc.Range('J18').Formula='=J16'
        for cell in ('G17','G19','G20'):calc.Range(cell).Value2=1
        calc.Range('J11').Value2=0;calc.Range('J14').Value2=0
        calc.Range('D3').Value2='';calc.Range('D4').Value2=''
        # A placeholder shear/moment is replaced with the visible connection links below.
        calc.Range('J12').Value2=0;calc.Range('J13').Value2=0
        for cell in ('L46','L48','L50'):calc.Range(cell).Formula='=J11'
        calc.Calculate()
        expected=bracket_actions(b,numeric(waler,'R15'),numeric(calc,'R15'))
        at=lambda c,o:f'{c}{row+o}'
        cref=lambda c:f"'{calc.Name}'!{c}"
        wref=lambda c:f"'{waler.Name}'!{c}"
        # Keep the template's load, shear and moment presentation, with consistent units.
        copy_rows(source,sheet,25,35,row)
        sheet.Range(at('B',0)).Value2=f"Bracket Check: {b['bracket_id']} / {b['case']}"
        sheet.Range(at('E',1)).Formula='='+wref('P11');sheet.Range(at('P',1)).Formula='='+cref('P11')
        sheet.Range(at('B',2)).Value2='Waling weight';sheet.Range(at('H',2)).Value2='kN'
        sheet.Range(at('H',3)).Value2='kN';sheet.Range(at('H',4)).Value2='kN'
        # Use editable project geometry beside the source block, not the waler buckling length.
        for offset in (1,2,3,4):
            sheet.Range(f'{at("U",offset)}:{at("AA",offset)}').UnMerge()
            sheet.Range(f'{at("U",offset)}:{at("AA",offset)}').Merge()
            sheet.Range(f'{at("AB",offset)}:{at("AH",offset)}').UnMerge()
            sheet.Range(f'{at("AB",offset)}:{at("AH",offset)}').Merge()
            sheet.Range(at('AB',offset)).NumberFormat='0.000'
            if offset in (1,2):sheet.Range(at('AB',offset)).Interior.Color=65535
        sheet.Range(at('U',1)).Value2='Bracket spacing (m)';sheet.Range(at('AB',1)).Value2=b['tributary_length_m']
        sheet.Range(at('U',2)).Value2='Bracket m';sheet.Range(at('AB',2)).Value2=b['length_m']
        sheet.Range(at('U',3)).Value2='Load model';sheet.Range(at('AB',3)).Value2=b['load_model']
        sheet.Range(at('U',4)).Value2='Load e = L/2 (m)';sheet.Range(at('AB',4)).Formula='='+at('AB',2)+'/2'
        # Reuse row 30 for the selfweight and actual actions. No next-page reference.
        sheet.Range(f'{at("B",5)}:{at("AH",5)}').UnMerge();sheet.Range(f'{at("B",5)}:{at("AH",5)}').Clear()
        sheet.Range(f'{at("B",5)}:{at("I",5)}').Merge();sheet.Range(at('B',5)).Value2='Bracket selfweight q (kN/m)'
        sheet.Range(at('B',3)).Value2='Root shear V';sheet.Range(at('B',4)).Value2='Root moment M'
        sheet.Range(f'{at("F",2)}:{at("S",4)}').UnMerge();sheet.Range(f'{at("F",2)}:{at("S",4)}').Clear()
        for offset,unit in ((2,'kN'),(3,'kN'),(4,'kNm')):
            sheet.Range(f'{at("G",offset)}:{at("L",offset)}').Merge()
            sheet.Range(at('G',offset)).NumberFormat='0.000'
            sheet.Range(at('M',offset)).Value2=unit
        sheet.Range(f'{at("K",5)}:{at("P",5)}').Merge();sheet.Range(at('K',5)).NumberFormat='0.000'
        formulas={at('G',2):f'=IF(AND(COUNT({at("AB",1)},{at("AB",2)})=2,MIN({at("AB",1)},{at("AB",2)})>0),1.4*77*{wref("R15")}/10000*{at("AB",1)},"INPUTS REQUIRED")',
                  at('K',5):f'=1.4*77*{cref("R15")}/10000',
                  at('G',3):f'={at("G",2)}+{at("K",5)}*{at("AB",2)}',
                  at('G',4):f'=IF(AND(COUNT({at("G",2)},{at("K",5)},{at("AB",2)})=3,OR({at("AB",3)}="uniform",AND({at("AB",3)}="point",COUNT({at("AB",4)})=1,{at("AB",4)}>=0,{at("AB",4)}<={at("AB",2)}))),IF({at("AB",3)}="point",{at("G",2)}*{at("AB",4)},{at("G",2)}*{at("AB",2)}/2)+{at("K",5)}*{at("AB",2)}^2/2,"INPUTS REQUIRED")'}
        # Template Bracket member calculation remains inspectable as a hidden dependency.
        calc.Range('J16').Formula=f"='{sheet.Name}'!{at('AB',2)}"
        calc.Range('J12').Formula=f"='{sheet.Name}'!{at('G',3)}";calc.Range('J13').Formula=f"='{sheet.Name}'!{at('G',4)}"
        sheet.Range(at('B',6)).Value2='Bracket - Sheet Pile: Direct Shear Weld'
        sheet.Range(at('B',8)).Value2=b['leg_mm'];sheet.Range(at('O',9)).Value2=b['strength_mpa']
        sheet.Range(at('U',9)).Value2=b['effective_length_mm']/2;sheet.Range(at('Z',9)).Value2=2
        formulas.update({at('F',7):'='+at('G',3),at('J',9):'='+at('B',8),
                         at('G',10):f'=0.7*{at("J",9)}*{at("O",9)}*{at("U",9)}*{at("Z",9)}/1000',at('K',10):'='+at('F',7)})
        cap,force,status=at('G',10),at('K',10),at('AE',10)
        formulas[status]=guarded_status(cap,force)
        formulas[at('J',10)]=f'=IF(COUNT({cap},{force})<2,"?",IF({cap}>={force},">=","<"))'
        for cell,formula in formulas.items():sheet.Range(cell).Formula=formula
        for cell in (cap,force,at('F',7)):sheet.Range(cell).NumberFormat='0.00';sheet.Range(cell).ShrinkToFit=True
        # The complete member check stays on the AA Bracket detail worksheet,
        # keeping the supplied connection page's three-block layout.
        calc.Range('B7:AH7').UnMerge();calc.Range('B7:AH7').Merge()
        calc.Range('B7').Value2='Bracket weld on Connection Checks: direct shear per template; weld-group moment transfer is not assessed.'
        calc.Range('B7').WrapText=True;calc.Rows(7).RowHeight=27
        wb.Application.CalculateFullRebuild()
        for cell,value in [(at('G',3),expected['shear_kn']),(at('G',4),expected['moment_knm'])]:
            if not math.isclose(numeric(sheet,cell),value,rel_tol=1e-9):raise ValueError('托架荷载与独立算式不一致')
        capacity=numeric(sheet,cap);expected_cap=.7*b['leg_mm']*b['strength_mpa']*b['effective_length_mm']/1000
        if not math.isclose(capacity,expected_cap,rel_tol=1e-9):raise ValueError('托架焊缝与独立算式不一致')
        checks={'Shear V/Pv':numeric(calc,'J12')/numeric(calc,'E29'),
                'Moment M/Mc':numeric(calc,'J13')/numeric(calc,'D36'),
                **{label:numeric(calc,cell) for label,cell in [('Section Eq.8.78','U39'),('Buckling Eq.8.79','U46'),('Buckling Eq.8.80','U48'),('Buckling Eq.8.81','U50')]}}
        if not all(finite(value) and value>=0 for value in checks.values()):raise ValueError('托架构件验算未返回有效利用率')
        calc.Visible=2
        results.append(dict(input=b,actions=expected,checks=checks,weld_capacity_kn=capacity,
                            weld_demand_kn=expected['shear_kn'],weld_scope='direct_shear_only',
                            status='PASS' if capacity>=expected['shear_kn'] and max(checks.values())<=1 else 'FAIL',
                            calculation_sheet=calc.Name,weld_capacity_cell=cap,weld_force_cell=force,weld_status_cell=status,
                            load_cells=dict(tributary_length=at('AB',1),length=at('AB',2),model=at('AB',3),eccentricity=at('AB',4))))
        row+=12
    return row,results
