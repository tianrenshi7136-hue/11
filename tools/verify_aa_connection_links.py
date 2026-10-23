"""Recalculate edits in a disposable AA result copy, never the original template."""
import json
import math
import shutil
from pathlib import Path

from els.sheetpile import excel_session
from els.supports import numeric


def main():
    root=Path(__file__).resolve().parent.parent
    result=json.loads((root/'outputs/aa-connections-result.json').read_text(encoding='utf8'))
    copy=root/'outputs/aa-connection-edit-test.xlsx';shutil.copyfile(result['workbook'],copy)
    b=result['support_package_results']['brackets'][0]
    with excel_session(copy) as (app,wb):
        c=wb.Worksheets('Connection Checks');calc=wb.Worksheets(b['calculation_sheet'])
        app.CalculateFullRebuild()
        before=numeric(calc,'J12');moment=numeric(calc,'J13');own=1.4*77*numeric(calc,'R15')/10000
        cells=b['load_cells'];c.Range(cells['tributary_length']).Value2=7
        app.CalculateFullRebuild()
        assert math.isclose(numeric(calc,'J12'),2*(before-own*.5)+own*.5,rel_tol=1e-9)
        c.Range(cells['eccentricity']).Value2=.25;app.CalculateFullRebuild()
        expected=(2*(before-own*.5))*.25+own*.5**2/2
        assert math.isclose(numeric(calc,'J13'),expected,rel_tol=1e-9)
        c.Range(cells['eccentricity']).MergeArea.ClearContents();app.CalculateFullRebuild()
        assert calc.Range('J13').Value2=='INPUTS REQUIRED'
        for cell,value in [(cells['eccentricity'],.5),(cells['tributary_length'],3.5)]:c.Range(cell).Value2=value
        app.CalculateFullRebuild()
        assert math.isclose(numeric(calc,'J13'),moment,rel_tol=1e-9)
        c.Range(b['weld_capacity_cell']).Value2=.01;app.CalculateFullRebuild()
        assert c.Range(b['weld_status_cell']).Value2=='NOT OK'
        summary=wb.Worksheets('Strut & Wailing Summary');summary.Range('R12').Value2=350;app.CalculateFullRebuild()
        sw=next(w for w in result['support_package_results']['welds'] if w['input']['kind']=='Strut-Waler')
        assert math.isclose(numeric(c,sw['force_cell']),1.4*350/3.5*4/2,rel_tol=1e-9)
        for name in result['output_sheets']:
            s=wb.Worksheets(name)
            assert s.PageSetup.CenterHorizontally and s.PageSetup.CenterVertically
    copy.unlink()
    print('PASS: bracket spacing, centroid, missing-input guards, weld failure and strut linkage')


if __name__=='__main__':main()
