# coding=utf-8
import xlrd, xlwt, re, datetime
from xlutils.copy import copy
from os import path, listdir

"""
get_tracking_file
    v1.0: 根据追踪数据文件名返回Workbook及可写入数据的行index。若无此文件，新建；若有，利用xlutils复制
get_date_time
    v1.0: 依次返回当前日期时间中的年、月、日
get_date
    v1.0: 返回文件名中对应的月和日，mmdd格式
get_filename
    v1.0: 根据港口名称关键词，返回结果文件和数据追踪文件的文件名
read_list
    v1.0: 分别返回主流粉矿、主流块矿、非主流资源、钢厂、贸易商的列表
    v1.1: 分别返回主流粉矿、主流块矿、非主流资源、品种、钢厂、贸易商的列表
    v1.2: 分别返回主流粉矿、主流块矿、非主流资源、品种、钢厂、贸易商、要追踪数据的粉矿、块矿的列表
    v1.3: 更改结构：第一个子表列出后面各个子表的题目。增加整个list的扩展性，可增加新的分类，每个分类中可增加具体项目。
read_data
    v1.0: 根据指定关键词判断标题行、货物列、货主列、数量列位置，根据指定条件判断是否为有效数据行，
          依次读取并返回货主、货物、数量的字典
    v1.1: 根据指定关键词判断标题行、货物列、货主列、数量列位置，根据指定条件判断是否为有效数据行，
          根据品种字典判断是否是需要读取的货物，依次读取并返回货主、货物、数量的字典
judge_merger_cell
    v1.0: 判断是否存在合并单元格，若有，输出合并单元格范围
read_merge_cell
    v1.0: 如果有合并单元格的情况，返回字典结果，key为序号，value为子表index、合并单元格的行index、列index、值
"""

def get_tracking_file(trackname):
    if path.exists(trackname.decode('utf-8')):
        originfile = xlrd.open_workbook(trackname.decode('utf-8'), 'r')
        trackfile = copy(originfile)
        subsheet = trackfile.get_sheet(0)
        rowindex = originfile.sheets()[0].nrows
        olddate = {}
        for i in range(5, rowindex):
            olddate[5] = originfile.sheets()[0].cell_value(i, 0)
    else:
        trackfile = xlwt.Workbook()
        subsheet = trackfile.add_sheet("Tracking Data")
        rowindex = 0
        olddate = {}
    return (trackfile, subsheet, rowindex, olddate)

def get_date_time():
    now_time = datetime.datetime.now()
    date = now_time.strftime('%Y/%m/%d')
    year = now_time.strftime('%Y')
    month = now_time.strftime('%m')
    day = now_time.strftime('%d')
    return (date, year, month, day)

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
        print u"文件名中日期格式不适合，请将日期统一为**.**格式."
    #print std_date
    return std_date

def get_filename (filename):
    stddate = get_date(filename)
    prefix = "铁矿港存结构分析-"
    namelist = ["岚桥", "岚山", "连云港", "京唐港", "实业", "青岛", "日照"]
    resultname = ''
    trackname = ''
    for item in namelist:
        if item in filename:
            resultname = prefix+item+"-"+stddate+".xls"
            trackname = item+"-历史追踪数据.xls"
    if resultname and trackname:
        return (resultname, trackname, stddate)
        #return resultname
    else:
        print u"未找到文件名港口关键词，请检查文件名."

def read_list(listname):
    """读取分类名录文件中的主流粉矿、主流块矿、非主流资源、品种、钢厂、贸易商
    分别为一个子表各自返回一个列表"""
    listfile = xlrd.open_workbook(listname.decode('utf-8'), 'r')
    class_list = listfile.sheets()[0].col_values(0)
    kinds = listfile.sheets()[1].col_values(0)
    company = listfile.sheets()[2].col_values(0)
    goods_class_name = {}
    goods_class_list = {}
    for i in range(2, len(class_list)):
        goods_class_name[i-2] = class_list[i]
        goods_class_list[i-2] = listfile.sheets()[i+1].col_values(0)
    #for item in class_list:
    #    print u'读取"%s"文件中的"%s"清单.' % (listname.decode('utf-8'), item)
    print u'已读取"%s"文件中的\033[1;34;0m%d\033[0m个清单.' % (listname.decode('utf-8'), len(class_list))
    return (kinds, company, goods_class_list, goods_class_name)

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
            elif title[k] in [u"结 存 量", u"库存", u"港存数", u"结存"] or u"数量" in title[k]:
                amountindex = k
        #print ownerindex, goodsindex, amountindex

        ### 遍历每行，检查是否是所需数据，是则读取入相应字典 ###
        uncountlist = []
        for k in range(titleindex+1, sheets[i-1].nrows):
            data = sheets[i-1].row_values(k)
            if data[goodsindex] and ((data[goodsindex] in kinds) or (u"精粉" in data[goodsindex]) or (u"球团" in data[goodsindex])) \
                    and (not (u"合计" in data[goodsindex])) and isinstance(data[amountindex],float) and (data[amountindex] > 0): #判断货物和数量列是否有数据，不是合计的数据，是在品种清单中的货物。
                count += 1
                owner[count] = data[ownerindex]
                goods[count] = data[goodsindex]
                amount[count] = data[amountindex]
            elif data[goodsindex] and (not ((data[goodsindex] in kinds) or (u"精粉" in data[goodsindex]) or (u"球团" in data[goodsindex]))) \
                    and (not (u"合计" in data[goodsindex])) and isinstance(data[amountindex],float):
                uncountlist.append(data[goodsindex])
        for item in set(uncountlist):
            print u'"\033[1;31;0m%s\033[0m"没有统计。若应统计在内，请检查分类名录中的“品种”清单是否包含.' % item

    ### 返回存储数据的3个字典 ###
    print u'共计读取\033[1;34;0m%d\033[0m条数据.' % count
    return (owner, goods, amount)

def judge_merger_cell(sheets, sheetindex):
    for i in sheetindex:
        sheet = sheets[i-1]
        mergedict = {}
        mergerange = sheet.merged_cells
        if mergerange:
            for k in range(0, len(mergerange)):
                #if (mergerange[k][3]-mergerange[k][2]) > 1:
                print "Mergerd cell found in subsheet %d: Column %s->%s, Row %d->%d" \
                      % (i, chr(65+mergerange[k][2]), chr(65+mergerange[k][3]-1), mergerange[k][0]+1, mergerange[k][1])

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

def getCustomFiles(prefix, path):
    allfiles = listdir(path)
    filelist = []
    for item in allfiles:
        if prefix in item.decode('gb2312'):
            #print item.decode('gb2312')
            filelist.append(item.decode('gb2312'))
    return filelist