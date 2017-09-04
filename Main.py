# coding=utf-8
import xlrd, xlwt, re
from Read_Data import *
from Write_Data import *

### 需要用户定义的变量 ###
filename = "京唐港8.1库存.xlsx"   # Excel数据文件的文件名，带扩展名。
sheetindex = [1]  # 需要读取的子表序号(第几个)，有多个时以英文逗号,间隔。

# 相对固定的设置，如有改动，需相应改变设置
listname = "分类名录.xlsx"  # 记录主流粉矿、主流块矿、非主流资源、品种、钢厂、贸易商名录的文件
stdname = "标准名称.xlsx"  # 记录货主（钢厂、贸易商）、品种标准名称的数据文件




#########################

### 确定结果文件名和对应日期 ###
resultname = get_filename(filename)
#resultname = "铁矿港存结构分析-岚桥-0804.xls"   # 用于自定义输出文件的文件名，或get_filename函数出错时使用

### 打开文件，读取数据 ###
mainpowder, mainblock, nonmain, kinds, company, trader = read_list(listname)
datafile = xlrd.open_workbook(filename.decode('utf-8'))
sheets = datafile.sheets()
judge_merger_cell(sheets, sheetindex)
owner, goods, amount = read_data(sheets, sheetindex, kinds)
# 打印输出测试
#for i in range(1, len(owner.keys())+1):
#    print '"%d": "%s", "%s", "%d"' % (i, owner[i], goods[i], amount[i])

### 货主、货物名称标准化 ###
standardize_name(stdname, owner, goods)

### 输出统计及详细数据 ###
resultfile = xlwt.Workbook()
write_summary(resultfile, mainpowder, mainblock, nonmain, owner, goods, amount, company, trader)
write_detail(resultfile, owner, goods, amount)
resultfile.save(resultname.decode('utf-8'))
print 'Summary and Detail Results Have Been Written in File "%s".' % resultname.decode('utf-8')
