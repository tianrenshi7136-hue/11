"""Native Excel verification of independent outputs and maximum-force selection."""
import copy
import json
import math
import shutil
import zipfile
from pathlib import Path
from els.supports import generate_supports
from els.sheetpile import excel_session
from tools.verify_aa_connections_excel import sample_package


def package():
    p=sample_package();p.pop('welds')
    s=copy.deepcopy(p['members'][0]);s.update(member_id='S2',case='LC2',section='203x203x86',anchor_kn=500,angle_deg=0)
    w=copy.deepcopy(p['members'][1]);w.update(member_id='W2',case='LC2',section='305x305x240',source_strut_id='S2',source_strut_case='LC2')
    p['members'] += [s,w]
    p['connection_groups']=[dict(kind='Strut-Waler',enabled=True,leg_mm=8,strength_mpa=220),dict(kind='Waler-Sheetpile',enabled=True,leg_mm=4,strength_mpa=220,effective_length_mm=800,distribution_mode='span')]
    return p


def main():
    import argparse
    parser=argparse.ArgumentParser();parser.add_argument('--existing',action='store_true');args=parser.parse_args()
    root=Path(__file__).resolve().parent.parent
    result=(json.loads((root/'outputs/v06-result.json').read_text(encoding='utf8')) if args.existing else generate_supports(package(),root/'templates/sheetpile-structure.local.json',root/'outputs/v06-validation','Software Validation Sample','Maximum Force Connection Checks'))
    checks=result['support_package_results'];assert len(checks['welds'])==2
    sw=next(w for w in checks['welds'] if w['input']['kind']=='Strut-Waler');ww=next(w for w in checks['welds'] if w['input']['kind']=='Waler-Sheetpile')
    assert sw['input']['member_id']=='S2' and ww['input']['member_id']=='W2'
    with zipfile.ZipFile(result['workbook']) as z:
        assert not any(name.startswith('xl/externalLinks/') for name in z.namelist())
    temp=root/'outputs/v06-edit-test.xlsx';shutil.copyfile(result['workbook'],temp)
    with excel_session(temp) as (app,wb):
        assert not wb.LinkSources(1)
        b=checks['brackets'][0];c=wb.Worksheets('Connection Checks');calc=wb.Worksheets(b['calculation_sheet']);cells=b['load_cells']
        assert c.Range(cells['eccentricity']).HasFormula
        c.Range(cells['length']).Value2=.8;app.CalculateFullRebuild()
        assert math.isclose(c.Range(cells['eccentricity']).Value2,.4,rel_tol=1e-9)
        p=1.4*77*wb.Worksheets('W-W1-LC1').Range('R15').Value2/10000*3.5
        q=1.4*77*calc.Range('R15').Value2/10000
        assert math.isclose(calc.Range('J12').Value2,p+q*.8,rel_tol=1e-9)
        assert math.isclose(calc.Range('J13').Value2,p*.4+q*.8**2/2,rel_tol=1e-9)
        summary=wb.Worksheets('Strut & Wailing Summary');summary.Range('R12').Value2=600;app.CalculateFullRebuild()
        c=wb.Worksheets('Connection Checks');force=c.Range(sw['force_cell']).Value2
        assert math.isclose(force,1.4*600/math.cos(math.radians(45)),rel_tol=1e-9)
        row=int(''.join(ch for ch in sw['force_cell'] if ch.isdigit()))-9
        assert 'S1 / LC1' in c.Range(f'B{row}').Value2
        summary.Range('R12').MergeArea.ClearContents();app.CalculateFullRebuild()
        assert c.Range(sw['status_cell']).Value2=='CHECK INPUTS'
    temp.unlink()
    (root/'outputs/v06-result.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf8')
    print(json.dumps({'status':'PASS','workbook':result['workbook'],'pdf':result['pdf'],'selected':['S2 / LC2','W2 / LC2'],'checks':'no external links; live maximum force; missing-input guard'},ensure_ascii=False))


if __name__=='__main__':main()
