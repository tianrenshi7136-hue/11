import copy,json,math,shutil
from pathlib import Path
from els.sheetpile import excel_session,generate
r=Path.cwd();result=json.loads((r/'outputs/supports-integration-result.json').read_text(encoding='utf8'));req=result['request'];temp=r/'outputs/support-edit-test.xls';shutil.copy2(result['workbook'],temp)
with excel_session(temp) as (app,wb):
 wb.Worksheets('Member Inputs').Range('M9').Value2=100
 app.CalculateFullRebuild()
 expected=1.4*100*3.5/math.cos(math.radians(45))
 assert math.isclose(wb.Worksheets('S-S1-LC1').Range('J11').Value2,expected)
 assert math.isclose(wb.Worksheets('Member Summary').Range('W9').Value2,expected)
 assert math.isclose(wb.Worksheets('Weld Checks').Range('K9').Value2,expected)
 wb.Worksheets('Blinding Concrete').Range('D11').Value2=100;app.CalculateFullRebuild()
 assert wb.Worksheets('Blinding Concrete').Range('H13').Value2=='FAIL'
temp.unlink()
for mode in ('weld','blinding','member'):
 bad=copy.deepcopy(req)
 if mode=='weld': bad['support_package']['welds'][0]['leg_mm']=1
 elif mode=='blinding':bad['support_package']['blinding_thickness_mm']=100
 else:bad['support_package']['members'][0]['anchor_kn_per_m']=10000
 out=r/'outputs'/('supports-failure-'+mode)
 try:generate(bad,r/'templates/sheetpile-structure.local.json',out,'Failure Test','ELS Checks')
 except ValueError as e:
  assert '未通过' in str(e)
 else: raise AssertionError(mode+' failure unexpectedly accepted')
 assert not list(out.rglob('*.xls')) and not list(out.rglob('*.pdf'))
print('PASS: edited inputs propagate; member/weld/blinding failures leave no XLS/PDF')
