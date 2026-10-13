import copy,json,math,shutil
from pathlib import Path
from els.sheetpile import excel_session,generate
r=Path.cwd();result=json.loads((r/'outputs/supports-integration-result.json').read_text(encoding='utf8'));req=result['request'];temp=r/'outputs/support-edit-test.xls';shutil.copy2(result['workbook'],temp)
with excel_session(temp) as (app,wb):
 wb.Worksheets('Strut & Wailing Summary').Range('R12').Value2=100
 app.CalculateFullRebuild()
 expected=1.4*100*3.5/math.cos(math.radians(45))
 assert math.isclose(wb.Worksheets('S-S1-LC1').Range('J11').Value2,expected)
 assert math.isclose(wb.Worksheets('Strut & Wailing Summary').Range('X12').Value2,expected)
 assert math.isclose(wb.Worksheets('Connection Checks').Range('O16').Value2,expected)
 visible=[s.Name for s in wb.Worksheets if s.Visible==-1]
 assert 'Member Data' not in visible and 'Member Inputs' not in visible
 for name in result['output_sheets']:
  sheet=wb.Worksheets(name)
  assert sheet.PageSetup.CenterHorizontally and sheet.PageSetup.CenterVertically
  end=sheet.Range(sheet.PageSetup.PrintArea).Rows.Count
  assert all(v in (None,'') for row in sheet.Range(f'A{end-1}:AH{end}').Value2 for v in row)
  assert sheet.Range(f'A1:AH{end}').Borders(9).Weight==4
  assert not sheet.Rows(end-1).Hidden and not sheet.Rows(end).Hidden
 connection=wb.Worksheets('Connection Checks')
 for row in (7,8,9,10,11,12,13,14,15,16,25,26,27,28,29,30,31,32,33,34):
  assert not connection.Rows(row).Hidden
 wb.Worksheets('Strut & Wailing Summary').Range('R12').MergeArea.ClearContents()
 app.CalculateFullRebuild()
 assert wb.Worksheets('S-S1-LC1').Range('J11').Value2=='INPUTS REQUIRED'
 assert connection.Range('AE16').Value2=='CHECK INPUTS'
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

