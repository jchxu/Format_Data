# coding=utf-8
import xlrd, xlwt, re
from Read_Data import *
from Write_Data import *

### 需要用户定义的变量 ###
whichdate = ''  # 是否指定汇总某特定日期的数据，例如0804
resultnameprefix = "铁矿港存结构分析汇总"


portfiles = getCustomFiles(u'铁矿港存结构分析-', r'.')
filedict = classbydate(portfiles)
dates, port, owner, goods, amount = get_sum_data(whichdate, filedict)

