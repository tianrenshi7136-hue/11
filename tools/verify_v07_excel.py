"""Native XLSX generation, including failing checks and all six FSP sections."""
import json,math,tempfile,shutil,zipfile
from pathlib import Path
from els.sheetpile import draft,generate,excel_session,check_template,calculate_candidate,envelope
from els.wall_compat import wall_reference
from els.sections import SECTIONS,properties
from tools.verify_aa_connections_excel import sample_package


def main():
 root=Path(__file__).resolve().parent.parent
 r=draft();r.update(section='FSP IV',yield_strength_mpa=275,shear_reduction=.7,bending_reduction=.7,stiffness_reduction=.5,analysis_ei_knm2_per_m=39565,source='SOFTWARE TEST ONLY',parameter_reference='SOFTWARE TEST ONLY',cases=[dict(name='ULS1',v_kn_per_m=1000,m_knm_per_m=2000)])
 with tempfile.TemporaryDirectory() as folder:
  path=Path(folder)/'sections.xlsx';shutil.copyfile(wall_reference(root/'templates/sheetpile-structure.local.json'),path)
  with excel_session(path) as (app,wb):
   check_template(wb)
   for section in SECTIONS:
    c=calculate_candidate(wb,section,r,envelope(r));assert math.isclose(c['ei_knm2_per_m'],properties(section,.5)['reduced_ei_knm2_per_m'],rel_tol=1e-9)
 p=sample_package();p.update(include_blinding=True,blinding_thickness_mm=100,length_m=10,width_m=5)
 p['members'][0]['anchor_kn']=5000
 for weld in p['welds']:weld['leg_mm']=.1
 r['support_package']=p
 result=generate(r,root/'templates/sheetpile-structure.local.json',root/'outputs/v07-validation','Software Validation Sample','Unified Generation - Failed Checks')
 assert result['check_status']=='NOT PASS' and result['selected']['status']=='NOT OK'
 checks=result['support_package_results'];assert checks['status']=='NOT PASS' and checks['failed_checks']
 assert checks['blinding']['status']=='FAIL'
 assert any(w['status']=='FAIL' for w in checks['welds'])
 with zipfile.ZipFile(result['workbook']) as z:assert not any(n.startswith('xl/externalLinks/') for n in z.namelist())
 (root/'outputs/v07-result.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf8')
 print(json.dumps({'status':'PASS','workbook':result['workbook'],'pdf':result['pdf'],'checks':checks['failed_checks'],'six_sections':'PASS'},ensure_ascii=False),flush=True)


if __name__=='__main__':main()
