"""Retain the engineer's original summary tables and connection calculation blocks."""
import math


def copy_rows(source,target,first,last,destination):
    source.Range(f'A{first}:AH{last}').Copy(target.Range(f'A{destination}'))
    for offset,row in enumerate(range(first,last+1)):
        target.Rows(destination+offset).Hidden=False
        target.Rows(destination+offset).RowHeight=source.Rows(row).RowHeight or 15


def set_formula(ws,cell,formula):
    ws.Range(cell).Formula=formula


def combined_status(member,row):
    ref=lambda c:f"'{member.Name}'!{c}"
    checks=','.join(ref(c) for c in ('U39','U46','U48','U50'))
    return f'=IF(COUNT({checks})<4,"CHECK INPUTS",IF(AND(MAX({checks})<=1,O{row}<=R{row},U{row}<=X{row},AA{row}<=AD{row}),"OK","NOT OK"))'


def build_summary(wb,summary,package,member_sheets):
    source=wb.Worksheets('Strut&Wailing Summary')
    summary.Range('A7:AH120').UnMerge();summary.Range('A7:AH120').Clear()
    copy_rows(source,summary,9,11,9)
    summary.Range('O10').Value2='Analysis output w (kN/m)'
    summary.Range('R10').Value2='Adopted w (kN/m)'
    force_kn=package.get('schema_version')==2
    if force_kn:
        for cell,text in [('I10','Strut spacing (m)'),('O10','Normal force F (kN)'),('R10','Adopted F (kN)'),('U10','Design 1.4F (kN)')]:summary.Range(cell).Value2=text
    index_by_key={(m['member_id'],m['case']):i for i,m in enumerate(package['members'],9)}
    struts=[m for m in package['members'] if m['kind']=='Strut']
    walers=[m for m in package['members'] if m['kind']=='Waler']
    for row,m in enumerate(struts,12):
        copy_rows(source,summary,12,12,row)
        member,design=member_sheets[(m['member_id'],m['case'])];i=index_by_key[(m['member_id'],m['case'])]
        data=lambda col:f"'Member Data'!{col}{i}"
        summary.Range(f'B{row}').Value2=m['member_id']+' / '+m['case']
        design['axial_summary_row']=row
        design['summary_sheet']=summary.Name
        for col,formula in [('D',f"='{member.Name}'!P11"),('I','='+data('P')),('L','='+data('V')),('O','='+data('M')),('R',f'=ABS(O{row})'),('U',f'=R{row}*1.4'),('X',f'=IF(AND(COUNT(I{row},L{row},R{row})=3,I{row}>0,L{row}>=0,L{row}<90),I{row}*U{row}/COS(RADIANS(L{row})),"INPUTS REQUIRED")'),('AC',f"='{member.Name}'!U36"),('AA',f'=IF(X{row}<=AC{row},"<=",">")'),('AF',f'=IF(COUNT(X{row},AC{row})<2,"CHECK INPUTS",IF(X{row}<=AC{row},"OK","NOT OK"))')]:
            set_formula(summary,f'{col}{row}',formula)
        member.Range('J11').Formula=f"='{summary.Name}'!X{row}"
        if force_kn:
            summary.Range(f'X{row}').Formula=f'=IF(AND(COUNT(I{row},L{row},R{row})=3,I{row}>0,L{row}>=0,L{row}<90),U{row}/COS(RADIANS(L{row})),"INPUTS REQUIRED")'
    heading=12+len(struts)+2
    copy_rows(source,summary,16,18,heading)
    summary.Range(f'AD{heading+1}').Value2='Capacity Mc (kNm)'
    start=heading+3
    def action_row(m,row,editable):
        copy_rows(source,summary,19,19,row)
        member,design=member_sheets[(m['member_id'],m['case'])];i=index_by_key[(m['member_id'],m['case'])]
        design['summary_row']=row
        data=lambda col:f"'Member Data'!{col}{i}"
        summary.Range(f'B{row}').Value2=m['member_id']+' / '+m['case']
        formulas={'D':f"='{member.Name}'!P11",'H':'='+data('M'),'K':'='+data('P'),'M':'='+data('S'),
                  'O':f'=IF(AND(COUNT(H{row},K{row},M{row})=3,MIN(K{row},M{row})>0),1.4*ABS(H{row})*K{row},"INPUTS REQUIRED")' if editable else f"='{member.Name}'!J11",
                  'R':f"='{member.Name}'!U35",'U':f'=1.4*ABS(H{row})*M{row}/2' if editable else f"='{member.Name}'!J12",
                  'X':f"='{member.Name}'!E30",'AA':f'=1.4*ABS(H{row})*M{row}^2/8' if editable else f"='{member.Name}'!J13",
                  'AD':f"='{member.Name}'!D36",'Q':f'=IF(O{row}<=R{row},"<=",">")','W':f'=IF(U{row}<=X{row},"<=",">")','AC':f'=IF(AA{row}<=AD{row},"<=",">")', 'AG':combined_status(member,row)}
        for col,formula in formulas.items():set_formula(summary,f'{col}{row}',formula)
        if editable:
            if force_kn:
                linked,_=member_sheets[(m['source_strut_id'],m['source_strut_case'])]
                linked_row=member_sheets[(m['source_strut_id'],m['source_strut_case'])][1]['axial_summary_row']
                summary.Range(f'H{row}').Formula=f"=ABS('{summary.Name}'!R{linked_row})/'{summary.Name}'!I{linked_row}"
                summary.Range(f'K{row}').Formula=f"='{summary.Name}'!I{linked_row}"
            for cell,col in [('J11','O'),('J12','U'),('J13','AA')]:member.Range(cell).Formula=f"='{summary.Name}'!{col}{row}"
        else:
            if force_kn:
                axial_row=design['axial_summary_row']
                summary.Range(f'H{row}').Formula=f"='{summary.Name}'!R{axial_row}"
                summary.Range(f'K{row}').Formula=f"='{summary.Name}'!I{axial_row}"
                if m.get('lx_is_default'):member.Range('J16').Formula=f"='{summary.Name}'!M{row}"
                if m.get('lx_equals_ly'):member.Range('J18').Formula='=J16'
            member.Range('J12').Formula=f"=1.4*77*R15/10000*'{summary.Name}'!M{row}/2"
            member.Range('J13').Formula=f"=J12*'{summary.Name}'!M{row}/4"
    for row,m in enumerate(walers,start):action_row(m,row,True)
    last=start+max(len(walers),1)-1
    if struts:
        heading=last+3;copy_rows(source,summary,16,18,heading)
        summary.Range(f'B{heading}').Value2='III. Strut Shear and Moment Checks'
        summary.Range(f'AD{heading+1}').Value2='Capacity Mc (kNm)'
        if force_kn:
            summary.Range(f'H{heading+1}').Value2='Normal F (kN)'
            summary.Range(f'K{heading+1}').Value2='Spacing (m)'
            summary.Range(f'M{heading+1}').Value2='Strut length (m)'
        for row,m in enumerate(struts,heading+3):action_row(m,row,False)
        last=heading+2+len(struts)
    for row in range(12,last+1):
        for col in ('H','I','L','K','M','O','R','U','X','AA','AC','AD'):
            summary.Range(f'{col}{row}').NumberFormat='0.00'
            summary.Range(f'{col}{row}').ShrinkToFit=True
    wb.Application.CalculateFullRebuild()
    for member,design in member_sheets.values():
        row=design['summary_row']
        design['summary_status']=summary.Range(f'AG{row}').Value2
        if design['summary_status']!='OK':
            design['status']='FAIL'
    return last


def build_welds(wb,sheet,package,member_sheets,numeric):
    source=wb.Worksheets('Connection check')
    sheet.Range('A7:AH200').UnMerge();sheet.Range('A7:AH200').Clear()
    row=7;results=[]
    for w in package['welds']:
        member,design=member_sheets[(w['member_id'],w['case'])]
        at=lambda col,offset:f'{col}{row+offset}'
        ref=lambda cell:f"'{member.Name}'!{cell}"
        label=w['member_id']+' / '+w['case']
        if w['kind']=='Strut-Waler':
            copy_rows(source,sheet,7,16,row)
            corner=design['member']['angle_deg']>0
            sheet.Range(at('B',0)).Value2=f'Connection Check: {label} ({"Corner Strut" if corner else "Strut"} - Waler)'
            sheet.Range(at('B',1)).Value2='Design Strut Force per Width'
            width=design['member']['load_width_m']
            sheet.Range(at('Q',2)).Value2=width
            sheet.Range(at('Q',2)).Formula=f"='{design['summary_sheet']}'!I{design['axial_summary_row']}"
            sheet.Range(at('Q',2)).NumberFormat='0.00'
            for col,offset in [('M',1),('M',2),('O',9)]:
                sheet.Range(at(col,offset)).NumberFormat='0.00'
            formulas={at('M',1):f'={ref("J11")}/{at("Q",2)}',at('M',2):f'={at("M",1)}',at('W',2):f'={at("M",2)}*{at("Q",2)}',
                      at('J',4):'='+ref('R17'),at('J',5):f'={ref("R17")}-{ref("R18")}',at('J',6):'='+ref('R20')}
            sheet.Range(at('D',3)).Value2=w['leg_mm'];sheet.Range(at('U',8)).Value2=w['strength_mpa']
            formulas[at('O',9)]=f'={at("W",2)}'
            cap_cell=at('J',9);force_cell=at('O',9);status_cell=at('AE',9)
            formulas[at('N',9)]=f'=IF({cap_cell}>={force_cell},">=","<")'
            length=2*(numeric(member,'R17')+(numeric(member,'R17')-numeric(member,'R18'))+numeric(member,'R20'))
            demand=design['design_actions']['n'];units='kN';count=10
        else:
            copy_rows(source,sheet,18,23,row)
            sheet.Range(at('B',0)).Value2=f'Connection Check: {label} (Waler - Sheet Piles)'
            sheet.Range(at('H',1)).Value2='kN from waler'
            sheet.Range(at('Q',1)).Value2=w['distribution_length_m']
            sheet.Range(f'{at("Q",1)}:{at("R",1)}').Merge()
            sheet.Range(at('S',1)).Value2='m long'
            sheet.Range(at('Q',1)).NumberFormat='0.00'
            sheet.Range(at('B',3)).Value2=w['leg_mm'];sheet.Range(at('O',4)).Value2=w['strength_mpa']
            sheet.Range(at('U',4)).Value2=w['effective_length_mm']/2
            formulas={at('F',1):'='+ref('J12'),at('F',2):f'={at("F",1)}/{at("Q",1)}',at('K',5):f'={at("F",2)}'}
            cap_cell=at('G',5);force_cell=at('K',5);status_cell=at('AE',5)
            formulas[at('J',5)]=f'=IF({cap_cell}>={force_cell},">=","<")'
            length=w['effective_length_mm'];demand=design['design_actions']['v']/w['distribution_length_m'];units='kN/m';count=6
        formulas[status_cell]=f'=IF(COUNT({cap_cell},{force_cell})<2,"CHECK INPUTS",IF(AND({cap_cell}>0,{cap_cell}>={force_cell}),"OK","NOT OK"))'
        for cell,formula in formulas.items():set_formula(sheet,cell,formula)
        sheet.Calculate();cap=numeric(sheet,cap_cell)
        expected=.7*w['leg_mm']*w['strength_mpa']*length/1000
        if not math.isclose(cap,expected,rel_tol=1e-9) or not math.isclose(numeric(sheet,force_cell),demand,rel_tol=1e-9):raise ValueError('Connection check与独立算式不一致')
        results.append({'input':w,'demand':demand,'capacity':cap,'units':units,'effective_length_mm':length,'status':'PASS' if cap>=demand else 'FAIL','capacity_cell':cap_cell,'force_cell':force_cell,'status_cell':status_cell})
        row+=count+1
    return results,row-2
