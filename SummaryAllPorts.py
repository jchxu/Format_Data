# coding=utf-8
import xlrd, xlwt, re
from Read_Data import *
from Write_Data import *

### 需要用户定义的变量 ###
whichdate = ''  # 是否指定汇总某特定日期的数据，例如0804
resultnameprefix = "铁矿港存结构分析汇总-"
trackname = "铁矿港存结构分析汇总历史追踪.xls"
listname = "分类名录.xlsx"

portfiles = getCustomFiles(u'铁矿港存结构分析-', r'.')
filedict = classbydate(portfiles)
kinds, company, goods_class_list, goods_class_name = read_list(listname)
dates, port, owner, goods, amount = get_all_data(whichdate, filedict)
totalrow, mainrow, classrow, goodsrow = calculate_sum_summary(dates, owner, goods, amount, company, goods_class_list, goods_class_name)

for item in dates.keys():
    resultfile = xlwt.Workbook()
    write_sum_summary(resultfile, item, goods_class_list, totalrow[item], mainrow[item], classrow[item], goodsrow[item])
    write_sum_detail(resultfile, item, dates, port, owner, goods, amount)
    resultname = resultnameprefix+get_date_time()[1]+item+".xls"
    resultfile.save(resultname.decode('utf-8'))

    trackfile, subsheet, rowindex, olddate = get_tracking_file(trackname)
    write_sum_tracking(item, trackfile, subsheet, rowindex, goods_class_name, goods_class_list, totalrow[item], mainrow[item], classrow[item], goodsrow[item])
    trackfile.save(trackname.decode('utf-8'))