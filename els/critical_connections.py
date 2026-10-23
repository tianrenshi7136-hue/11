"""One connection check per family, governed by maximum design action."""
import copy

from .core import finite


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
        for field in ('leg_mm','strength_mpa')+(('effective_length_mm',) if target=='Waler' else ()):
            if not finite(group.get(field)) or group[field]<=0:errors.append(f'{kind} {field}须填写正数')
        if target=='Waler' and group.get('distribution_mode') not in ('span','manual'):errors.append('请选择围檩焊缝的荷载分配长度方式')
    return errors


def governing_welds(package,member_sheets):
    grouped={};result=[]
    for weld in package.get('welds',[]):
        if not weld.get('group_id'):result.append(weld);continue
        grouped.setdefault(weld['group_id'],[]).append(weld)
    for candidates in grouped.values():
        def force(weld):
            key=(weld['member_id'],weld['case'])
            return member_sheets[key][1]['design_actions']['n' if weld['kind']=='Strut-Waler' else 'v']
        chosen=copy.deepcopy(max(candidates,key=force))
        chosen['_candidates']=candidates;chosen['_selection_basis']='maximum_design_axial_force' if chosen['kind']=='Strut-Waler' else 'maximum_design_shear'
        result.append(chosen)
    return result


def attach_live_selector(wb,sheet,weld,member_sheets,row):
    """Excel MAX/INDEX keeps the selected case live after editing member inputs."""
    candidates=weld.get('_candidates')
    if not candidates:return
    try:data=wb.Worksheets('Connection Candidates')
    except Exception:
        data=wb.Worksheets.Add();data.Name='Connection Candidates'
        data.Range('A1:H1').Value2=(('Member','Case','Section','Design N or V (kN)','B (mm)','t (mm)','d (mm)','Distribution (m)'),)
    first=max(3,data.UsedRange.Row+data.UsedRange.Rows.Count+2)
    for index,candidate in enumerate(candidates,first):
        member,design=member_sheets[(candidate['member_id'],candidate['case'])]
        data.Cells(index,1).Value2=candidate['member_id'];data.Cells(index,2).Value2=candidate['case']
        for col,cell in ((3,'P11'),(4,'J11' if weld['kind']=='Strut-Waler' else 'J12'),(5,'R17'),(6,'R18'),(7,'R20')):
            data.Cells(index,col).Formula=f"='{member.Name}'!{cell}"
        if candidate.get('distribution_mode')=='span':data.Cells(index,8).Formula=f"='{design['summary_sheet']}'!M{design['summary_row']}"
        else:data.Cells(index,8).Value2=candidate.get('distribution_length_m')
    last=first+len(candidates)-1
    data.Range(f'J{first}').Formula=f'=IF(COUNT(D{first}:D{last})={len(candidates)},MATCH(MAX(D{first}:D{last}),D{first}:D{last},0),0)'
    def pick(col):return f'IF(\'Connection Candidates\'!$J${first}>0,INDEX(\'Connection Candidates\'!{col}{first}:{col}{last},\'Connection Candidates\'!$J${first}),"INPUTS REQUIRED")'
    label='Connection Check: Strut / Corner Strut - Waling (' if weld['kind']=='Strut-Waler' else 'Connection Check: Waling - Sheet Pile ('
    sheet.Range(f'B{row}').Formula=f'="{label}"&{pick("A")}&" / "&{pick("B")}&")"'
    if weld['kind']=='Strut-Waler':
        for offset in (1,2):sheet.Range(f'M{row+offset}').Formula='='+pick('D')
        sheet.Range(f'J{row+4}').Formula='='+pick('E')
        sheet.Range(f'J{row+5}').Formula='='+pick('E')+'-'+pick('F')
        sheet.Range(f'J{row+6}').Formula='='+pick('G')
    else:
        sheet.Range(f'F{row+1}').Formula='='+pick('D')
        sheet.Range(f'Q{row+1}').Formula='='+pick('H')
    data.Visible=2
