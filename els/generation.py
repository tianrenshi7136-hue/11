"""Select complete calculation modules without changing the user's draft."""
import copy
from .core import finite
from .sheetpile import validate_request


def generation_plan(steel,package):
    steel=copy.deepcopy(steel);p=copy.deepcopy(package);missing=[]
    positive=lambda v:finite(v) and v>0
    steel_ready=bool(steel) and not validate_request(steel)
    include_steel=p.get('include_sheetpile',True)
    if include_steel and not steel_ready:missing.append('钢板桩')
    members=[]
    for kind,label in [('Strut','Struts 支撑'),('Waler','Waling 围檩')]:
        rows=[m for m in p['members'] if m['kind']==kind]
        fields=('strut_length_m','spacing_m','lx_m','ly_m','kx','ky','klt') if kind=='Strut' else ('load_span_m','lx_m','ly_m','kx','ky','klt')
        ready=bool(rows) and all(m.get('member_id') and m.get('case') and m.get('section') not in (None,'','请选择') and m.get('grade') in ('S235','S275','S355','S460') and all(positive(m.get(k)) for k in fields) and (finite(m.get('anchor_kn')) and finite(m.get('angle_deg')) if kind=='Strut' else (m.get('source_strut_id'),m.get('source_strut_case')) in {(s['member_id'],s['case']) for s in members}) for m in rows)
        if ready:members.extend(rows)
        else:missing.append(label)
    p['members']=members
    walers={(m['member_id'],m['case']) for m in members if m['kind']=='Waler'}
    brackets=p.get('brackets',[])
    if not brackets or not all((b.get('waler_id'),b.get('waler_case')) in walers and b.get('section') not in (None,'','请选择') and b.get('grade') in ('S235','S275','S355','S460') and all(positive(b.get(k)) for k in ('length_m','tributary_length_m','leg_mm','strength_mpa','effective_length_mm')) for b in brackets):
        missing.append('Bracket 托架');p['brackets']=[]
    groups=[]
    for group in p.get('connection_groups',[]):
        if not group.get('enabled'):continue
        kind='Strut' if group['kind']=='Strut-Waler' else 'Waler'
        required=('leg_mm','strength_mpa')+(('effective_length_mm',) if kind=='Waler' else ())+(('distribution_length_m',) if kind=='Waler' and group.get('distribution_mode')=='manual' else ())
        has_force_source=kind!='Strut' or any(m['kind']=='Waler' for m in members)
        if has_force_source and any(m['kind']==kind for m in members) and all(positive(group.get(k)) for k in required):groups.append(group)
    if not groups:missing.append('Weld 连接')
    elif len(groups)<sum(bool(g.get('enabled')) for g in p.get('connection_groups',[])):missing.append('部分 Weld 连接')
    p['connection_groups']=groups;p['welds']=[]
    blinding=p.get('include_blinding') and positive(p.get('blinding_thickness_mm')) and positive(p.get('length_m')) and positive(p.get('width_m')) and steel_ready
    if not blinding and (p.get('include_blinding') or not p.get('blinding_thickness_mm')):missing.append('Blinding 垫层')
    p['include_blinding']=bool(blinding)
    return dict(steel=steel if include_steel and steel_ready else None,support=p if members or blinding else None,blinding_steel=steel if blinding else None,missing=missing)
