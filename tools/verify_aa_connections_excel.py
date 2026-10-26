"""Software test inputs only. Requires the registered local engineer templates."""
import json
from pathlib import Path

from els.connections import AA_CONNECTIONS
from els.supports import generate_supports


def sample_package():
    return dict(schema_version=2,connection_template=AA_CONNECTIONS,enabled=True,
                include_blinding=False,source='SOFTWARE TEST ONLY',
                members=[dict(member_id='S1',case='LC1',kind='Strut',section='152x152x37',grade='S355',
                              anchor_kn=315,spacing_m=3.5,strut_length_m=2.5,angle_deg=45,
                              lx_m=2.5,ly_m=2.5,lx_is_default=True,lx_equals_ly=True,kx=1,ky=1,klt=1),
                         dict(member_id='W1',case='LC1',kind='Waler',section='203x203x86',grade='S355',
                              source_strut_id='S1',source_strut_case='LC1',load_span_m=4,
                              lx_m=3.5,ly_m=3.5,kx=1,ky=1,klt=1)],
                welds=[dict(member_id='S1',case='LC1',kind='Strut-Waler',leg_mm=8,strength_mpa=220),
                       dict(member_id='W1',case='LC1',kind='Waler-Sheetpile',leg_mm=4,strength_mpa=220,
                            effective_length_mm=800,distribution_length_m=3.5)],
                brackets=[dict(bracket_id='B1',case='LC1',waler_id='W1',waler_case='LC1',
                               section='102x102x23',grade='S355',length_m=.5,tributary_length_m=3.5,
                               load_model='point',eccentricity_m=.5,leg_mm=4,strength_mpa=220,effective_length_mm=200)])


def main():
    root=Path(__file__).resolve().parent.parent
    result=generate_supports(sample_package(),root/'templates/sheetpile-structure.local.json',
                             root/'outputs/aa-connections-validation','Software Validation Sample',
                             'AA Template Connection Checks')
    (root/'outputs/aa-connections-result.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf8')
    print(json.dumps({'workbook':result['workbook'],'pdf':result['pdf'],'checks':result['support_package_results']},ensure_ascii=False,indent=2))


if __name__=='__main__':main()
