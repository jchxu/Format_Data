# coding=utf-8
import xlrd, xlwt, re
from Read_Data import *
from Write_Data import *

### 需要用户定义的变量 ###
filename = "连云港贸易矿8.4.xls"   # Excel数据文件的文件名，带扩展名。
sheetindex = [1]  # 需要读取的子表序号(第几个)，有多个时以英文逗号,间隔。

# 相对固定的设置，如有改动，需相应改变设置
listname = "分类名录.xlsx"  # 记录主流粉矿、主流块矿、非主流资源、品种、钢厂、贸易商名录的文件
stdname = "标准名称.xlsx"  # 记录货主（钢厂、贸易商）、品种标准名称的数据文件
tracklist = "追踪品种.xlsx"  # 记录需追踪数据品种的文件名
#########################

### 确定结果文件名和对应日期 ###
resultname, trackname, stddate = get_filename(filename)
#resultname = "铁矿港存结构分析-岚桥-0804.xls"   # 用于自定义输出文件的文件名，或get_filename函数出错时使用

### 打开文件，读取数据 ###
kinds, company, goods_class_list, goods_class_name = read_list(listname)
#mainpowder, mainblock, nonmain, kinds, company, trader, powder, block = read_list(listname)
datafile = xlrd.open_workbook(filename.decode('utf-8'))
sheets = datafile.sheets()
judge_merger_cell(sheets, sheetindex)
owner, goods, amount = read_data(sheets, sheetindex, kinds)
### 打印输出测试
#for item in goodslistname.values():
#    print "%s" % item
#for i in range(1, len(owner.keys())+1):
#    print '"%d": "%s", "%s", "%d"' % (i, owner[i], goods[i], amount[i])

### 货主、货物名称标准化 ###
standardize_name(stdname, owner, goods)

### 输出统计及详细数据 ###
resultfile = xlwt.Workbook()
totalrow, mainrow, classrow, goodsrow = calculate_summary(owner, goods, amount, company, goods_class_list, goods_class_name)
print u'统计及详细信息将写入"\033[1;34;0m%s\033[0m"文件.' % resultname.decode('utf-8')
write_summary(resultfile, goods_class_list, totalrow, mainrow, classrow, goodsrow)
write_detail(resultfile, owner, goods, amount)
resultfile.save(resultname.decode('utf-8'))


### 输出历史追踪数据 ###
trackfile, subsheet, rowindex, olddate = get_tracking_file(trackname)
#goodsdata = calculate_trackdata(powder, block, goodsrow, owner, goods, amount, company)
trackfile, writeindex = write_tracking(tracklist, stddate, olddate, trackfile, subsheet, rowindex, goods_class_name, goods_class_list, totalrow, mainrow, classrow, goodsrow)
trackfile.save(trackname.decode('utf-8'))
print u'历史追踪数据已写入"\033[1;34;0m%s\033[0m"文件，第\033[1;34;0m%d\033[0m行.' % (trackname.decode('utf-8'), writeindex+1)