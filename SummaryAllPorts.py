# coding=utf-8
import xlrd, xlwt, re
from Read_Data import *
from Write_Data import *

### 需要用户定义的变量 ###
whichdate = ''  # 是否指定汇总某特定日期的数据，例如0804
resultnameprefix = "铁矿港存结构分析汇总-"
trackname = "铁矿港存结构分析汇总历史追踪.xls"
listname = "分类名录.xlsx"  # 记录主流粉矿、主流块矿、非主流资源、品种、钢厂、贸易商名录的文件
stdname = "标准名称.xlsx"  # 记录货主（钢厂、贸易商）、品种标准名称的数据文件


portfiles = getCustomFiles(u'铁矿港存结构分析-', r'.')
filedict = classbydate(portfiles)
kinds, company, trader, goods_class_list, goods_class_name = read_list(listname)
dates, port, owner, goods, amount = get_all_data(whichdate, filedict)
standardize_name(stdname, owner, goods)
totalrow, mainrow, classrow, goodsrow = calculate_sum_summary(dates, owner, goods, amount, company, goods_class_list, goods_class_name)

### 各港口汇总统计，写入Summary、Detail、历史追踪数据 ###
for item in dates.keys():
    resultfile = xlwt.Workbook()
    write_sum_summary(resultfile, item, goods_class_list, totalrow[item], mainrow[item], classrow[item], goodsrow[item])
    write_sum_detail(resultfile, item, dates, port, owner, goods, amount)
    resultname = resultnameprefix+get_date_time()[1]+item+".xls"
    resultfile.save(resultname.decode('utf-8'))

    trackfile, subsheet, rowindex, olddate = get_tracking_file(trackname)
    write_sum_tracking(item, trackfile, subsheet, rowindex, goods_class_name, goods_class_list, totalrow[item], mainrow[item], classrow[item], goodsrow[item])
    trackfile.save(trackname.decode('utf-8'))

### 按照贸易商、品种统计货权集中度并排序 ###
