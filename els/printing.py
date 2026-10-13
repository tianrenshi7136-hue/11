"""Shared engineering calculation print layout."""


def setup_print(ws, last_content_row, app):
    bottom = last_content_row + 2
    ws.Range(f'A{last_content_row}:AH{last_content_row}').Borders(9).Weight = 2
    padding = ws.Range(f'A{last_content_row+1}:AH{bottom}')
    padding.UnMerge()
    padding.Clear()
    ws.Rows(f'{last_content_row+1}:{bottom}').Hidden = False
    ws.Rows(f'{last_content_row+1}:{bottom}').RowHeight = 15
    ws.Rows(f'{bottom+1}:65536').Hidden = True
    ws.PageSetup.PrintArea = f'$A$1:$AH${bottom}'
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
        border=ws.Range(f'A1:AH{bottom}').Borders(edge)
        border.LineStyle,border.Weight=1,4
