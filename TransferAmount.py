# coding=utf-8
import xlrd, xlwt, re
from xlutils.copy import copy
from Write_Data import *
from Read_Data import *

filename = "连云港贸易矿8.4.xls"  # Excel数据文件的文件名，带扩展名。
sheetindex = 3  # 需要读取的子表序号(第几个子表？)。
colindex = 7  # 非标准数量所在列序号（第几列？）

### 读取子表，转换数量单位 ###
datafile = xlrd.open_workbook(filename.decode('utf-8'), 'r')
sheets = datafile.sheets()
subsheet = sheets[sheetindex-1]
amount = transfer_nick_amount(subsheet, colindex)
#print amount

### 重写数据文件 ###
resfile = copy(datafile)
subsheet2 = resfile.get_sheet(sheetindex-1)
subsheet2.write(min(amount.keys())-1, subsheet.ncols, u"数量")
for i in amount.keys():
    subsheet2.write(i, subsheet.ncols, amount[i])
resfile.save(filename.decode('utf-8'))
#resfile.save("test.xls")