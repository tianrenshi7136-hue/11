"""Reissue the annotated example with waling V and verify live dependencies."""
import copy,json,math,shutil,uuid
from pathlib import Path
from els.sheetpile import generate,excel_session,save_and_export
from els.critical_connections import attach_live_selector


def main():
 import argparse
 parser=argparse.ArgumentParser();parser.add_argument('--existing',action='store_true');parser.add_argument('--validate-only',action='store_true');parser.add_argument('--result',type=Path);args=parser.parse_args()
 root=Path(__file__).resolve().parent.parent
 old=json.loads((root/'outputs/v07-validation/20261023-143621-4c300a/result.json').read_text(encoding='utf8'))
 if args.result:
  r=json.loads(args.result.read_text(encoding='utf8'));args.validate_only=True
 elif args.existing or args.validate_only:
  result_file=max((root/'outputs/waling-weld-correction').glob('*/result.json'),key=lambda p:p.stat().st_mtime)
  r=json.loads(result_file.read_text(encoding='utf8'))
 if args.existing:
  with excel_session(r['workbook']) as (app,wb):
   member_sheets={(m['member']['member_id'],m['member']['case']):(wb.Worksheets(m['sheet']),m) for m in r['support_package_results']['members']}
   for w in r['support_package_results']['welds']:
    row=int(''.join(ch for ch in w['force_cell'] if ch.isdigit()))-(9 if w['input']['kind']=='Strut-Waler' else 5)
    attach_live_selector(wb,wb.Worksheets('Connection Checks'),w['input'],member_sheets,row)
   app.CalculateFullRebuild();save_and_export(wb,r['pdf'])
 elif not args.validate_only:r=generate(old['request'],root/'templates/sheetpile-structure.local.json',root/'outputs/waling-weld-correction',old['project_name'],old['subject_name'])
 sw=next(w for w in r['support_package_results']['welds'] if w['input']['kind']=='Strut-Waler')
 assert sw['input']['force_member_id']=='W1' and math.isclose(sw['demand'],4000,rel_tol=1e-10)
 temp=Path(r['workbook']).with_name('disposable-link-test-'+uuid.uuid4().hex[:8]+'.xlsx');shutil.copyfile(r['workbook'],temp)
 with excel_session(temp) as (app,wb):
  c=wb.Worksheets('Connection Checks');w1=wb.Worksheets('W-W1-LC1');s1=wb.Worksheets('S-S1-LC1')
  assert math.isclose(c.Range(sw['force_cell']).Value2,4000,rel_tol=1e-10)
  assert math.isclose(s1.Range('J11').Value2,9899.494936611665,rel_tol=1e-10)
  s1.Range('J11').Value2=1;app.CalculateFullRebuild()
  assert math.isclose(c.Range(sw['force_cell']).Value2,4000,rel_tol=1e-10)
  w1.Range('J12').Value2=4321;app.CalculateFullRebuild()
  assert c.Range(sw['force_cell']).Value2==4321
  w2=wb.Worksheets.Add(c);w2.Name='W-Test2';s2=wb.Worksheets.Add(c);s2.Name='S-Test2'
  for cell in ('P11','R17','R18','R20'):s2.Range(cell).Value2=s1.Range(cell).Value2
  w2.Range('J12').Value2=5000;s2.Range('J11').Value2=.1
  candidate=copy.deepcopy(sw['input']);candidate.update(member_id='S2',case='LC1',force_member_id='W2',force_case='LC1')
  test=dict(sw['input'],_candidates=[sw['input'],candidate])
  row=int(''.join(ch for ch in sw['force_cell'] if ch.isdigit()))-9
  attach_live_selector(wb,c,test,{('S1','LC1'):(s1,{}),('S2','LC1'):(s2,{}),('W1','LC1'):(w1,{}),('W2','LC1'):(w2,{})},row)
  app.CalculateFullRebuild();assert c.Range(sw['force_cell']).Value2==5000;assert 'W2' in c.Range(f'B{row}').Value2
  w2.Range('J12').Value2=3000;app.CalculateFullRebuild();assert c.Range(sw['force_cell']).Value2==4321
  w1.Range('J12').MergeArea.ClearContents();app.CalculateFullRebuild();assert c.Range(sw['status_cell']).Value2=='CHECK INPUTS'
 temp.unlink()
 (root/'outputs/waling-weld-correction-result.json').write_text(json.dumps(r,ensure_ascii=False,indent=2),encoding='utf8')
 print(json.dumps({'status':'PASS','waling_shear_kn':sw['demand'],'workbook':r['workbook'],'pdf':r['pdf'],'checks':'waling V only; no repeated 1.4; maximum waling V changes controlling pair; missing-input guard'},ensure_ascii=False),flush=True)


if __name__=='__main__':main()
