# coding=utf-8
import xlrd, xlwt, re
from xlutils.copy import copy
from Write_Data import *
from Read_Data import *

filename = "连云港贸易矿8.4.xls"  # Excel数据文件的文件名，带扩展名。

### 读取各子表中可能存在的合并单元格信息 ###
datafile = xlrd.open_workbook(filename.decode('utf-8'), 'r', formatting_info=True)
sheets = datafile.sheets()
mergedict, count = read_merge_cell(sheets)

### 若有合并的单元格，将其各单元格以相同内容填充 ###
if count > 0:
    resfile = copy(datafile)
    for i in range(0, count):
        sheet = resfile.get_sheet(mergedict[i][0])
        sheet.write(mergedict[i][1], mergedict[i][2], mergedict[i][3])
    resfile.save(filename.decode('utf-8'))
