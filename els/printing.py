"""Shared engineering calculation print layout."""


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
    for row in range(7,last_content_row+1):
        cell = ws.Cells(row,34)
        area = cell.MergeArea
        if (area.Column+area.Columns.Count-1 == 34 and
                area.Cells(1,1).Formula not in ('',None) and
                cell.Borders(10).LineStyle == 1):
            table_right_edges.add(area.Address)
        if full_separator:
            separator_cell = ws.Cells(row,29)
            separator_area = separator_cell.MergeArea
            if (separator_area.Columns.Count > 1 and
                    separator_area.Column+separator_area.Columns.Count-1 == 29 and
                    separator_area.Cells(1,1).Formula not in ('',None) and
                    separator_cell.Borders(10).LineStyle == 1):
                separator_table_edges.add(separator_area.Address)
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
    ws.Rows(f'{bottom+1}:65536').Hidden = True
    ws.PageSetup.PrintArea = f'$A$1:$AI${bottom}'
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = 1
    ws.PageSetup.CenterHorizontally = True
    ws.PageSetup.CenterVertically = True
    for margin in ('LeftMargin','RightMargin','TopMargin','BottomMargin'):
        setattr(ws.PageSetup,margin,app.InchesToPoints(.25))
    ws.PageSetup.HeaderMargin = app.InchesToPoints(.1)
    ws.PageSetup.FooterMargin = app.InchesToPoints(.1)
    for edge in (7,8,9,10):
        border=ws.Range(f'A1:AI{bottom}').Borders(edge)
        border.LineStyle,border.Weight=1,4
    if full_separator:
        divider = ws.Range(f'AD3:AD{bottom}').Borders(10)
        divider.LineStyle,divider.Weight = 1,2
