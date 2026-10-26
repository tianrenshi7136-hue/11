"""Run the same local seven-page example; compare values and printed formulas."""
import json,math,time
import zipfile,xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries
from els.sheetpile import generate


def equivalent(a,b,path='result'):
 if isinstance(a,dict):
  assert a.keys()==b.keys(),path
  for key in a:equivalent(a[key],b[key],path+'.'+key)
 elif isinstance(a,list):
  assert len(a)==len(b),path
  for i,(x,y) in enumerate(zip(a,b)):equivalent(x,y,f'{path}[{i}]')
 elif isinstance(a,(float,int)) and not isinstance(a,bool):
  assert math.isclose(a,b,rel_tol=1e-10,abs_tol=1e-9),(path,a,b)
 else:assert a==b,(path,a,b)


def printed_cells(path,names,data_only):
 book=load_workbook(path,read_only=True,data_only=data_only,keep_links=False)
 try:
  with zipfile.ZipFile(path) as z:
   xml=ET.fromstring(z.read('xl/workbook.xml'));ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
   areas={book.sheetnames[int(n.attrib['localSheetId'])]:n.text for n in xml.findall('m:definedNames/m:definedName',ns) if n.attrib['name']=='_xlnm.Print_Area'}
  result={}
  for name in names:
   ws=book[name];area=areas[name]
   # Excel writes one print range for each output sheet.
   bounds=range_boundaries(str(area).split('!')[-1].strip("'"))
   left,top,right,bottom=bounds
   result[name]=[[c.value for c in row] for row in ws.iter_rows(min_row=top,max_row=bottom,min_col=left,max_col=right)]
   if data_only:
    # The report date is intentionally volatile; all engineering values must
    # still match when a benchmark is rerun on a later calendar day.
    assert isinstance(result[name][3][16],datetime),(name,'Missing report date')
    result[name][3][16]='CURRENT REPORT DATE'
  return result
 finally:book.close()


def main():
 import argparse
 p=argparse.ArgumentParser();p.add_argument('--baseline',default='outputs/waling-weld-correction/20261023-172651-eb1729/result.json');p.add_argument('--result');p.add_argument('--profile',action='store_true');p.add_argument('--supports-only',action='store_true');args=p.parse_args()
 root=Path(__file__).resolve().parent.parent
 baseline=json.loads((root/args.baseline).read_text(encoding='utf8'))
 started=time.perf_counter()
 if args.result:
  result=json.loads((root/args.result).read_text(encoding='utf8'));elapsed=result['performance']['total_seconds']
 else:
  def run():
   if args.supports_only:
    from els.supports import generate_supports
    return generate_supports(baseline['request']['support_package'],root/'templates/sheetpile-structure.local.json',root/'outputs/performance-validation-supports',baseline['project_name'],baseline['subject_name'],baseline['request'],progress=lambda s:print(s,flush=True))
   return generate(baseline['request'],root/'templates/sheetpile-structure.local.json',root/'outputs/performance-validation',baseline['project_name'],baseline['subject_name'],progress=lambda s:print(s,flush=True))
  if args.profile:
   import cProfile
   profiler=cProfile.Profile();result=profiler.runcall(run);profiler.dump_stats(root/'outputs/performance-validation/profile.pstats')
  else:result=run()
  elapsed=time.perf_counter()-started
 def semantic(r):return {k:v for k,v in r.items() if k not in ('workbook','pdf','performance')}
 if args.supports_only:equivalent(baseline['support_package_results'],result['support_package_results'])
 else:equivalent(semantic(baseline),semantic(result))
 for data_only in (False,True):
  equivalent(printed_cells(baseline['workbook'],result['output_sheets'],data_only),printed_cells(result['workbook'],result['output_sheets'],data_only),'printed_cache' if data_only else 'printed_formulas')
 check=dict(status='PASS',elapsed_seconds=round(elapsed,3),baseline_profile_seconds=143.195,output_sheets=result['output_sheets'],workbook=result['workbook'],pdf=result['pdf'],performance=result['performance'],verified='same request; all result values; all printed formulas and cached values')
 (root/('outputs/performance-validation-supports-result.json' if args.supports_only else 'outputs/performance-validation-result.json')).write_text(json.dumps(check,ensure_ascii=False,indent=2),encoding='utf8')
 print(json.dumps(check,ensure_ascii=False),flush=True)


if __name__=='__main__':main()
