import json
from pathlib import Path
from els.sheetpile import generate
r=Path.cwd()
members=[]; welds=[]
for i,w in enumerate((90,130),1):
 for kind,section,span in [('Strut','152x152x37',2.5),('Waler','203x203x86' if i==1 else '254x254x132',4)]:
  mid=('S' if kind=='Strut' else 'W')+str(i)
  members.append(dict(member_id=mid,case='LC1',kind=kind,section=section,grade='S355',anchor_kn_per_m=w,load_width_m=3.5,load_span_m=span,angle_deg=45 if kind=='Strut' else 0,lx_m=2.5 if kind=='Strut' else 3.5,ly_m=2.5 if kind=='Strut' else 3.5,kx=1,ky=1,klt=1))
  welds.append(dict(member_id=mid,case='LC1',kind='Strut-Waler' if kind=='Strut' else 'Waler-Sheetpile',leg_mm=8 if kind=='Strut' else 4,strength_mpa=220,effective_length_mm=800,distribution_length_m=3.5))
req=dict(selection_mode='manual',section='FSP IV',yield_strength_mpa=275,shear_reduction=.7,bending_reduction=.7,stiffness_reduction=.5,analysis_ei_knm2_per_m=39565,source='SOFTWARE VALIDATION ONLY',parameter_reference='TEMPLATE BASELINE TEST ONLY',cases=[dict(name='LC1',v_kn_per_m=80.4,m_knm_per_m=74.1)],support_package=dict(source='TEMPLATE BASELINE TEST ONLY',members=members,welds=welds,length_m=10,width_m=5,blinding_thickness_mm=300))
result=generate(req,r/'templates/sheetpile-structure.local.json',r/'outputs/supports-validation','Software Validation Sample','ELS Structural Checks')
(r/'outputs/supports-integration-result.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'workbook':result['workbook'],'pdf':result['pdf'],'sheets':result['output_sheets'],'results':result['support_package_results']},ensure_ascii=False,indent=2))
