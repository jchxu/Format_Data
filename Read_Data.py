# coding=utf-8
import xlrd, xlwt, re, datetime
from xlutils.copy import copy
from os import path

"""
get_tracking_file
get_date_time
get_date
get_filename
read_list
    v1.0: 分别返回主流粉矿、主流块矿、非主流资源、钢厂、贸易商的列表
    v1.1: 分别返回主流粉矿、主流块矿、非主流资源、品种、钢厂、贸易商的列表
read_data
    v1.0: 根据指定关键词判断标题行、货物列、货主列、数量列位置，根据指定条件判断是否为有效数据行，
          依次读取并返回货主、货物、数量的字典
    v1.1: 根据指定关键词判断标题行、货物列、货主列、数量列位置，根据指定条件判断是否为有效数据行，
          根据品种字典判断是否是需要读取的货物，依次读取并返回货主、货物、数量的字典
read_merge_cell
"""

def get_tracking_file(trackname):
    if path.exists(trackname.decode('utf-8')):
        originfile = xlrd.open_workbook(trackname.decode('utf-8'), 'r')
        trackfile = copy(originfile)
        rowindex = originfile.sheets()[0].nrows
    else:
        trackfile = xlwt.Workbook()
        rowindex = 0
    return (trackfile, rowindex)

def get_date_time():
    now_time = datetime.datetime.now()
    year = now_time.strftime('%Y')
    month = now_time.strftime('%m')
    day = now_time.strftime('%d')
    return (year, month, day)

def get_date(filename):
    flag = filename.count('.') -1
    std_date = ''
    #print flag
    if flag == 0:
        date = re.findall(r"\d+", filename)[0]
        std_date = date
    elif flag == 1:
        date = re.findall(r"\d+\.?\d*", filename)[0]
        month = int(date.split('.')[0])
        day = int(date.split('.')[1])
        std_date = "%02d%02d" % (month, day)
    elif flag == 2:
        date = re.findall(r"\d+\.?\d+\.?\d*", filename)[0]
        month = int(date.split('.')[-2])
        day = int(date.split('.')[-1])
        std_date = "%02d%02d" % (month, day)
    else:
        print "文件名中日期格式不适合，请将日期统一为**.**格式。"
    #print std_date
    return std_date

def get_filename (filename):
    stddate = get_date(filename)
    prefix = "铁矿港存结构分析-"
    namelist = ["岚桥", "岚山", "连云港", "京唐港", "实业"]
    resultname = ''
    trackname = ''
    for item in namelist:
        if item in filename:
            resultname = prefix+item+"-"+stddate+".xls"
            trackname = item+"-历史追踪数据.xls"
    if resultname and trackname:
        #return (resultname, trackname)
        return resultname
    else:
        print "未找到文件名港口关键词，请检查文件名。"

def read_list(listname):
    """读取分类名录文件中的主流粉矿、主流块矿、非主流资源、品种、钢厂、贸易商
    分别为一个子表各自返回一个列表"""
    listfile = xlrd.open_workbook(listname.decode('utf-8'), 'r')
    mainpowder = listfile.sheets()[0].col_values(0)
    mainblock = listfile.sheets()[1].col_values(0)
    nonmain = listfile.sheets()[2].col_values(0)
    kinds = listfile.sheets()[3].col_values(0)
    company = listfile.sheets()[4].col_values(0)
    trader = listfile.sheets()[5].col_values(0)
    return (mainpowder, mainblock, nonmain, kinds, company, trader)

def read_data(sheets, sheetindex, kinds):
    """读取子表中的数据行。利用指定的特殊名词判断标题行、货物列、货主列、数量列位置
    根据指定条件判断是否为有效数据行，然后依次读取货主、货物、数量，最后返回一个字典。"""
    owner = {}
    goods = {}
    amount = {}
    count = 0

    for i in sheetindex:
        titleindex = 0
        ownerindex = 0
        goodsindex = 0
        amountindex = 0

        ### 确定标题行位置 ###
        for j in range(0, sheets[i-1].nrows):
            line = sheets[i-1].row_values(j)
            if u"船名" in line or u"船舶名称" in line or u"货主" in line:
                titleindex = j
                break
        title = sheets[i-1].row_values(titleindex)

        ### 确定货物、货主、数量列位置 ###
        for k in range(0, len(title)):
            if title[k] in [u"货种", u"货名", u"货性", u"品种"]:
                goodsindex = k
            elif title[k] in [u"货主"] or (u"收货人" in title[k]) or (u"钢厂" in title[k]):
                ownerindex = k
            elif title[k] in [u"结 存 量", u"库存", u"港存数"] or u"数量" in title[k]:
                amountindex = k
        #print ownerindex, goodsindex, amountindex

        ### 遍历每行，检查是否是所需数据，是则读取入相应字典 ###
        for k in range(titleindex+1, sheets[i-1].nrows):
            data = sheets[i-1].row_values(k)
            if data[goodsindex] and (data[goodsindex] in kinds) and (not (u"合计" in data[goodsindex])) \
                    and isinstance(data[amountindex],float): #判断货物和数量列是否有数据，不是合计的数据，是在品种清单中的货物。
                count += 1
                owner[count] = data[ownerindex]
                goods[count] = data[goodsindex]
                amount[count] = data[amountindex]
    ### 返回存储数据的3个字典 ###
    print "A total of %d records have been read." % count
    return (owner, goods, amount)

def read_merge_cell(sheets):
    mergedict = {}
    count = 0
    for i in range(0, len(sheets)):
        mergerange = sheets[i].merged_cells
        if mergerange:
            for k in range(0, len(mergerange)):
                rlow, rhigh, clow, chigh = mergerange[k]
                for m in range(rlow, rhigh):
                    for n in range(clow, chigh):
                        mergedict[count] = [i, m, n, sheets[i].cell_value(rlow, clow)]
                        count += 1
    print "A total of %d mergerd cells found." % count
    return (mergedict, count)