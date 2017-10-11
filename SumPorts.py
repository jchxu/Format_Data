# coding=utf-8
import xlrd, xlwt, re
from Read_Data import *
from Write_Data import *

portfiles = getCustomFiles(u'铁矿港存结构分析', r'.')

dates = []
date_files = {}
for item in portfiles:
    filename = item.strip('.xls').split('-')
    dates.append(filename[2].encode('utf-8'))
print dates
print list(set(dates))
for item in list(set(dates)):
    lists = []
    for portfile in portfiles:
        filename = portfile.strip('.xls').split('-')
        if item == filename[2].encode('utf-8'):
            lists.append(portfile)
    date_files[item] = lists
print date_files