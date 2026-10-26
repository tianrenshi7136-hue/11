"""One connection check per family, governed by maximum design action."""
import copy

from .core import finite


def associated_walers(package,weld):
    """A connection uses its waler's V and its linked strut's weld geometry."""
    return [m for m in package.get('members',[]) if m.get('kind')=='Waler' and
            (m.get('source_strut_id'),m.get('source_strut_case'))==(weld.get('member_id'),weld.get('case'))]


def force_key(weld):
    return weld.get('force_member_id',weld['member_id']),weld.get('force_case',weld['case'])


def expand_groups(package):
    if 'connection_groups' not in package:return
    welds=[]
    for group in package['connection_groups']:
        if not group.get('enabled'):continue
        member_kind='Strut' if group.get('kind')=='Strut-Waler' else 'Waler'
        for member in package.get('members',[]):
            if member.get('kind')!=member_kind:continue
            weld=dict(group,member_id=member['member_id'],case=member['case'],group_id=group['kind'])
            if group.get('distribution_mode')=='span':weld['distribution_length_m']=member.get('load_span_m')
            welds.append(weld)
    package['welds']=welds


def validate_groups(package):
    groups=package.get('connection_groups',[])
    if not isinstance(groups,list):return ['连接设置须为列表']
    errors=[];seen=set()
    for group in groups:
        if not isinstance(group,dict):errors.append('连接设置格式无效');continue
        kind=group.get('kind')
        if kind not in ('Strut-Waler','Waler-Sheetpile') or kind in seen:errors.append('连接类型无效或重复')
        seen.add(kind)
        if not group.get('enabled'):continue
        target='Strut' if kind=='Strut-Waler' else 'Waler'
        if not any(m.get('kind')==target for m in package.get('members',[])):errors.append('已勾选的连接类型没有对应构件')
        if target=='Strut' and not any(associated_walers(package,dict(member_id=m.get('member_id'),case=m.get('case'))) for m in package.get('members',[]) if m.get('kind')=='Strut'):
            errors.append('支撑—围檩焊缝须有已关联支撑的围檩，以读取围檩设计剪力')
        for field in ('leg_mm','strength_mpa')+(('effective_length_mm',) if target=='Waler' else ()):
            if not finite(group.get(field)) or group[field]<=0:errors.append(f'{kind} {field}须填写正数')
        if target=='Waler' and group.get('distribution_mode') not in ('span','manual'):errors.append('请选择围檩焊缝的荷载分配长度方式')
    return errors


def governing_welds(package,member_sheets):
    grouped={};result=[]
    for weld in package.get('welds',[]):
        if weld['kind']=='Strut-Waler':
            candidates=[dict(weld,force_member_id=m['member_id'],force_case=m['case']) for m in associated_walers(package,weld)]
            if not candidates and not weld.get('group_id'):raise ValueError('支撑—围檩连接缺少对应围檩，不能读取围檩剪力')
        else:candidates=[weld]
        if not weld.get('group_id') and weld['kind']!='Strut-Waler':result.append(weld);continue
        group=weld.get('group_id') or (weld['member_id'],weld['case'],weld['kind'])
        grouped.setdefault(group,[]).extend(candidates)
    for candidates in grouped.values():
        if not candidates:continue
        def force(weld):
            return abs(member_sheets[force_key(weld)][1]['design_actions']['v'])
        chosen=copy.deepcopy(max(candidates,key=force))
        chosen['_candidates']=candidates;chosen['_selection_basis']='maximum_waling_design_shear'
        result.append(chosen)
    return result


def attach_live_selector(wb,sheet,weld,member_sheets,row):
    """Excel MAX/INDEX keeps the selected case live after editing member inputs."""
    candidates=weld.get('_candidates')
    if not candidates:return
    try:data=wb.Worksheets('Connection Candidates')
    except Exception:
        data=wb.Worksheets.Add();data.Name='Connection Candidates'
        data.Range('A1:I1').Value2=(('Strut / Member','Case','Section','Waling design V (kN)','B (mm)','t (mm)','d (mm)','Distribution (m)','Waling'),)
    first=max(3,data.UsedRange.Row+data.UsedRange.Rows.Count+2)
    for index,candidate in enumerate(candidates,first):
        member,design=member_sheets[(candidate['member_id'],candidate['case'])]
        data.Cells(index,1).Value2=candidate['member_id'];data.Cells(index,2).Value2=candidate['case']
        for col,cell in ((3,'P11'),(5,'R17'),(6,'R18'),(7,'R20')):
            data.Cells(index,col).Formula=f"='{member.Name}'!{cell}"
        force_member,_=member_sheets[force_key(candidate)]
        reference=f"'{force_member.Name}'!J12"
        data.Cells(index,4).Formula=f'=IF(COUNT({reference})=1,ABS({reference}),"INPUTS REQUIRED")'
        data.Cells(index,9).Value2=force_key(candidate)[0]
        if candidate.get('distribution_mode')=='span':data.Cells(index,8).Formula=f"='{design['summary_sheet']}'!M{design['summary_row']}"
        else:data.Cells(index,8).Value2=candidate.get('distribution_length_m')
    last=first+len(candidates)-1
    data.Range(f'J{first}').Formula=f'=IF(COUNT(D{first}:D{last})={len(candidates)},MATCH(MAX(D{first}:D{last}),D{first}:D{last},0),0)'
    def pick(col):return f'IF(\'Connection Candidates\'!$J${first}>0,INDEX(\'Connection Candidates\'!{col}{first}:{col}{last},\'Connection Candidates\'!$J${first}),"INPUTS REQUIRED")'
    label='Connection Check: Strut / Corner Strut - Waling (' if weld['kind']=='Strut-Waler' else 'Connection Check: Waling - Sheet Pile ('
    sheet.Range(f'B{row}').Formula=f'="{label}"&{pick("A")}&" / "&{pick("B")}&")"'+(f'&"; Waling "&{pick("I")}' if weld['kind']=='Strut-Waler' else '')
    if weld['kind']=='Strut-Waler':
        for offset in (1,2):sheet.Range(f'M{row+offset}').Formula='='+pick('D')
        sheet.Range(f'J{row+4}').Formula='='+pick('E')
        sheet.Range(f'J{row+5}').Formula='='+pick('E')+'-'+pick('F')
        sheet.Range(f'J{row+6}').Formula='='+pick('G')
    else:
        sheet.Range(f'F{row+1}').Formula='='+pick('D')
        sheet.Range(f'Q{row+1}').Formula='='+pick('H')
    data.Visible=2
