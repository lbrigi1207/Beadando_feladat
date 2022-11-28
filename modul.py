import pathlib
from openpyxl import Workbook

def fajl():
    fajl = pathlib.Path('Adatok.xlsx')
    if fajl.exists():
        pass
    else:
        fajl = Workbook()
        sheet = fajl.active
        sheet['A1'] = 'Szerző neve'
        sheet['B1'] = 'Könyv címe'
        sheet['C1'] = 'Könyv hossza'
        sheet['D1'] = 'Könyv nyelve'
        sheet['E1'] = 'Ráfordított idő'
        sheet['F1'] = 'Értékelés'
        sheet['G1'] = 'Leírás'
        fajl.save('Adatok.xlsx')
