"""Native XLSX route for the same continuous-wall equations and FSP library."""
from pathlib import Path
from .reference_store import reference_file
from .sections import SECTIONS


def wall_reference(mapping_path):
    # The registered BIFF/XLS file is rejected by desktop Excel before any
    # formulas run. The engineer's AA workbook has the same wall method in
    # native XLSX; keep its layout and normalize the input/output row addresses.
    registry=Path(mapping_path).parent/'supports-structure.local.json'
    if registry.is_file():return reference_file(registry)
    return reference_file(mapping_path,role='wall')


def prepare_wall(wb):
    summary=wb.Worksheets('FSPiv Summary')
    if summary.Range('J7').Value2 not in SECTIONS:return
    for name in ('FSPiv Summary','Stiffness for FREW-FSPIV'):
        ws=wb.Worksheets(name);ws.Rows(7).Insert()
        if name=='FSPiv Summary':
            ws.Rows(28).Insert();ws.Rows(46).Insert()
        ws.Range('J8').Validation.Delete()
        ws.Range('J8').Validation.Add(3,1,1,'=$AK$9:$AK$14')
        for i,section in enumerate(SECTIONS,9):ws.Range(f'AK{i}').Value2=section
        ws.Range('J8').Value2='FSP IV'
    from .sheetpile import STATUS_FORMULA
    summary.Range('AB50').Formula=STATUS_FORMULA
    summary.Range('V33').Formula="='Stiffness for FREW-FSPIV'!V35"
    stiffness=wb.Worksheets('Stiffness for FREW-FSPIV')
    stiffness.Range('P35').Formula="='FSPiv Summary'!P33"
    # AA's native type and condition lists are unrelated to the engineering
    # section selector. Only the continuous wall is included in this module.
