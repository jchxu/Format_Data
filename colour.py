#! coding=utf-8
from xlrd import open_workbook
from xlwt import easyxf, XFStyle, Pattern
from xlutils.copy import copy

def set_style(colour):
    return easyxf('pattern: pattern solid, fore_colour %s;' % colour)

rbook = open_workbook('xlwt_colour.xls')
rsheet = rbook.sheet_by_index(0)
wbook = copy(rbook)
wsheet = wbook.get_sheet(0)

rows = rsheet.nrows
for r in range(rows):
    coltxt = str(rsheet.cell(r, 0).value)
    wsheet.write(r, 1, None, set_style(coltxt))

wsheet.col(0).width = 6000
wsheet.col(1).width = 3000

wbook.save('xlwt_colour.xls')