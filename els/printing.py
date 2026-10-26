"""Shared engineering calculation print layout."""
from openpyxl.utils.cell import range_boundaries
from .excel_ops import batched_print, hide_after


def format_check_results(ws):
    """Failing result text must never retain a template's green pass fill."""
    area=ws.Range(ws.PageSetup.PrintArea)
    first=area.Cells(1,1).Address.replace('$','')
    condition=area.FormatConditions.Add(2,1,f'=IFERROR(OR({first}="NOT OK",{first}="NOT PASS",{first}="FAIL",{first}="N.G."),FALSE)')
    condition.Interior.Color=14803455
    condition.Font.Color=128;condition.Font.Bold=True
    condition.SetFirstPriority()


@batched_print
def setup_print(ws, last_content_row, app):
    bottom = last_content_row + 2
    # AH belongs to the original tables. AI is the new outer frame column.
    if any(row[0] not in ('',None) for row in ws.Range(f'AI1:AI{bottom}').Formula):
        raise ValueError(f'{ws.Name}: AI列含数据，不能作为新增外框空白列')
    ws.Columns('AI').ColumnWidth = ws.Columns('AH').ColumnWidth
    ws.Range(f'AI1:AI{bottom}').ClearFormats()
    for index in range(ws.Shapes.Count,0,-1):
        if ws.Shapes.Item(index).TopLeftCell.Column > 34:
            ws.Shapes.Item(index).Delete()
    full_separator = (ws.Range('AC7').Borders(10).LineStyle == 1 or
                      ws.Range('AD7').Borders(7).LineStyle == 1)
    table_right_edges = set()
    separator_table_edges = set()
    formulas=ws.Range(f'A1:AH{last_content_row}').Formula
    def table_edges(column,number,merged_only=False):
        edges=set();row=7
        while row<=last_content_row:
            cell=ws.Range(f'{column}{row}')
            address=cell.MergeArea.Address
            left,top,right,bottom_row=range_boundaries(address)
            value=formulas[top-1][left-1]
            if (right==number and (not merged_only or right>left)
                    and value not in ('',None) and cell.Borders(10).LineStyle==1):
                edges.add(address)
            row=bottom_row+1
        return edges
    table_right_edges=table_edges('AH',34)
    if full_separator:separator_table_edges=table_edges('AC',29,True)
    ws.Range(f'AH1:AH{bottom}').Borders(10).LineStyle = -4142
    for address in table_right_edges:
        border = ws.Range(address).Borders(10)
        border.LineStyle,border.Weight = 1,2

    # Move the job-number header and its separator one column to the right.
    ws.Range('AD3:AH6').Copy(ws.Range('AE3'))
    ws.Range('AD3:AD6').Clear()
    for row in range(3,7):
        original = ws.Cells(row,29).Borders(9)
        if original.LineStyle == 1:
            extended = ws.Cells(row,30).Borders(9)
            extended.LineStyle,extended.Weight = original.LineStyle,original.Weight
    ws.Range('AC1:AC2').Borders(10).LineStyle = -4142
    ws.Range('AD1:AD2').Borders(7).LineStyle = -4142
    separator_bottom = bottom if full_separator else 6
    ws.Range(f'AC3:AC{separator_bottom}').Borders(10).LineStyle = -4142
    ws.Range(f'AD3:AD{separator_bottom}').Borders(7).LineStyle = -4142
    divider = ws.Range(f'AD3:AD{separator_bottom}').Borders(10)
    divider.LineStyle,divider.Weight = 1,2
    for address in separator_table_edges:
        border = ws.Range(address).Borders(10)
        border.LineStyle,border.Weight = 1,2

    title = ws.Range('A2:AI2')
    title.UnMerge()
    title.Merge()
    title.HorizontalAlignment = -4108
    title.Borders(9).LineStyle = 1
    title.Borders(9).Weight = 2
    # Preserve source table bottoms; never add a page-wide inner bottom line.
    padding = ws.Range(f'A{last_content_row+1}:AI{bottom}')
    padding.UnMerge()
    padding.Clear()
    ws.Rows(f'{last_content_row+1}:{bottom}').Hidden = False
    ws.Rows(f'{last_content_row+1}:{bottom}').RowHeight = 15
    hide_after(ws,bottom)
    page=ws.PageSetup
    page.PrintArea = f'$A$1:$AI${bottom}'
    page.Zoom = False
    page.FitToPagesWide = 1
    page.FitToPagesTall = 1
    page.CenterHorizontally = True
    page.CenterVertically = True
    for margin in ('LeftMargin','RightMargin','TopMargin','BottomMargin'):
        setattr(page,margin,18.0)
    page.HeaderMargin = 7.2
    page.FooterMargin = 7.2
    for edge in (7,8,9,10):
        border=ws.Range(f'A1:AI{bottom}').Borders(edge)
        border.LineStyle,border.Weight=1,4
    if full_separator:
        divider = ws.Range(f'AD3:AD{bottom}').Borders(10)
        divider.LineStyle,divider.Weight = 1,2
