"""Read BIFF formulas through a separate Excel instance, without saving."""
import json
import shutil
from pathlib import Path
import win32com.client
from els.templates import sha256

root = Path(__file__).resolve().parent.parent
app = win32com.client.DispatchEx('Excel.Application')
reports = []
try:
    app.Visible = False
    app.DisplayAlerts = False
    app.EnableEvents = False
    app.AutomationSecurity = 3
    blank = app.Workbooks.Add()
    app.Calculation = -4135
    for p in (root / 'TWD-TD004/design').rglob('*.xls'):
        before = sha256(p)
        folder = root / 'outputs/twd-inspection-copies'
        folder.mkdir(exist_ok=True)
        copy = folder / p.name
        shutil.copy2(p, copy)
        if sha256(copy) != before:
            raise RuntimeError('Source changed during snapshot copy')
        wb = app.Workbooks.Open(str(copy.resolve()), UpdateLinks=0, ReadOnly=True, IgnoreReadOnlyRecommended=True, AddToMru=False)
        try:
            sheets = []
            for ws in wb.Worksheets:
                formulas = []
                used = ws.UsedRange
                values = used.Formula
                import xlrd
                if isinstance(values, tuple):
                    for row, items in enumerate(values):
                        for col, value in enumerate(items):
                            if isinstance(value, str) and value.startswith('='):
                                formulas.append({'cell': xlrd.formula.colname(used.Column + col - 1) + str(used.Row + row), 'formula': value})
                elif isinstance(values, str) and values.startswith('='):
                    formulas.append({'cell': used.Address.replace('$', ''), 'formula': values})
                sample = {}
                if ws.Name in ('FSPiv Summary', 'Stiffness for FREW-FSPIV'):
                    for address in ('J7','J8','J9','J11','J12','J13','R18','R19','H39','H43','I47','T47','AB47','I48','T48','AB48','I50','AB50','R31','T31','S38','Q42','H44'):
                        c = ws.Range(address)
                        sample[address] = {'value': c.Value2, 'formula': c.Formula}
                sheets.append({'name': ws.Name, 'formulas': formulas, 'sample': sample})
            links = wb.LinkSources(1)
            reports.append({'path': str(p.resolve()), 'sha256': before, 'links': list(links or []), 'sheets': sheets})
        finally:
            wb.Close(SaveChanges=False)
        reports[-1]['source_stable_during_inspection'] = sha256(p) == before
        (root / 'outputs/twd-xls-formulas.json').write_text(json.dumps(reports, ensure_ascii=False, indent=2), encoding='utf-8')
        print(p.name, 'snapshot read', flush=True)
finally:
    app.Quit()
(root / 'outputs/twd-xls-formulas.json').write_text(json.dumps(reports, ensure_ascii=False, indent=2), encoding='utf-8')
