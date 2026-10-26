import json,math,shutil
from pathlib import Path
from els.supports import migrate_package,generate_supports
from els.sheetpile import excel_session
r=Path.cwd();old=json.loads((r/'outputs/supports-integration-result.json').read_text(encoding='utf8'))
p=migrate_package(old['request']['support_package']);p.update(enabled=True,include_blinding=False)
for m in p['members']:
 if m['kind']=='Strut':m.update(lx_is_default=True,lx_equals_ly=True)
result=generate_supports(p,r/'templates/sheetpile-structure.local.json',r/'outputs/supports-v2-validation','Software Validation Sample','Normal Force Strut and Waling Checks')
(r/'outputs/supports-v2-integration-result.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf8')
temp=r/'outputs/supports-v2-edit-copy.xls';shutil.copy2(result['workbook'],temp)
with excel_session(temp) as (app,wb):
 summary=wb.Worksheets('Strut & Wailing Summary')
 assert summary.Range('R12').Value2==315
 assert math.isclose(summary.Range('X12').Value2,1.4*315/math.cos(math.radians(45)))
 assert summary.Range('H19').Value2==90 and summary.Range('K19').Value2==3.5
 summary.Range('R12').Value2=350;app.CalculateFullRebuild()
 assert summary.Range('H19').Value2==100
 assert math.isclose(wb.Worksheets('S-S1-LC1').Range('J11').Value2,1.4*350/math.cos(math.radians(45)))
 assert summary.Range('U19').Value2==280
 summary.Range('I12').Value2=7;app.CalculateFullRebuild()
 assert summary.Range('K19').Value2==7 and summary.Range('H19').Value2==50
 assert math.isclose(wb.Worksheets('S-S1-LC1').Range('J11').Value2,1.4*350/math.cos(math.radians(45)))
 length_row=next(m for m in result['support_package_results']['members'] if m['sheet']=='S-S1-LC1')['summary_row']
 summary.Range(f'M{length_row}').Value2=3;app.CalculateFullRebuild()
 assert wb.Worksheets('S-S1-LC1').Range('J16').Value2==3 and wb.Worksheets('S-S1-LC1').Range('J18').Value2==3
temp.unlink()
print(json.dumps({'status':'PASS','workbook':result['workbook'],'pdf':result['pdf'],'pages':len(result['output_sheets'])},ensure_ascii=False))
