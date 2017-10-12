# coding=utf-8
import xlrd, xlwt, re
from Read_Data import *
from Write_Data import *

whichdate = ''

portfiles = getCustomFiles(u'铁矿港存结构分析', r'.')
filedict = classbydate(portfiles)
port, owner, goods, amount = get_sum_date(whichdate, filedict)

