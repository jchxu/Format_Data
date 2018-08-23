# coding=utf-8
import xlrd,re
from datetime import datetime

### 获取港口文件名中的简化港口名称,返回字典{港口index:简化名称} ###
def GetShortName(PortFilesList,StdPort):
    Shortnames = {}
    for i in range(0,len(PortFilesList)):
        PortFilename = PortFilesList[i]
        for item in StdPort:
            if item in PortFilename:
                Shortnames[i] = item
                break
    return Shortnames

### 获取各个港口文件名中的日期信息，返回字典{港口index：8位数日期}###
def GetFileDate(PortFilesList,StdPort):
    FileDateDict = {}
    ThisYear = str(datetime.now().year)
    for i in range(0,len(PortFilesList)):
        PortFilename = PortFilesList[i]
        Flag = PortFilename.count('.') - 1
        if Flag == 0:
            FileDate = re.findall(r"\d+", PortFilename)[0]
            if (len(FileDate) == 4): FileDate = ThisYear + FileDate
            elif (len(FileDate) == 6): FileDate = ThisYear[0:2] + FileDate
            FileDateDict[i] = FileDate
        elif Flag == 1:
            TempDate = re.findall(r"\d+", PortFilename)
            FileDate = ThisYear + "%02d"%int(TempDate[0]) + "%02d"%int(TempDate[1])
            FileDateDict[i] = FileDate
        elif Flag == 2:
            TempDate = re.findall(r"\d+", PortFilename)
            if (len(TempDate[0]) == 2): TempDate[0] = ThisYear[0:2] + TempDate[0]
            FileDate = TempDate[0] + "%02d" % int(TempDate[1]) + "%02d" % int(TempDate[2])
            FileDateDict[i] = FileDate
    return (FileDateDict)

### 获取文件名日期 ###
def GetFilename(filename):
    prefix = "铁矿港存结构分析-"
    namelist = ["岚桥", "岚山", "连云港", "京唐港", "实业", "青岛", "日照"]
    flag = filename.count('.') - 1
    std_date = ''
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
        print(u"文件名中日期格式不适合，请将日期统一为**.**格式.")
    for item in namelist:
        if item in filename:
            resultname = prefix+item+"-"+std_date
    return resultname

### 读取原始数据，返回列表 ###
def ReadSource(SourceFileName):
    Owner = []
    Goods = []
    Amount = []
    Port = []
    ArrivalDate = []
    SourceFile = xlrd.open_workbook(SourceFileName, 'r')
    Sheet = SourceFile.sheet_by_index(0)
    for i in range(1, Sheet.nrows):
        Line = Sheet.row_values(i)
        Owner.append(Line[0].replace(' ',''))
        Goods.append(Line[1].replace(' ',''))
        Amount.append(Line[2])
        Port.append(Line[3].replace(' ',''))
        if len(Line) >= 5:  #有日期，记录日期；无日期，记录为“-”
            ArrivalDate.append(Line[4])
        else:
            ArrivalDate.append('-')
    print(u'已读取"\033[1;34;0m%s\033[0m"中的\033[1;34;0m%d\033[0m条数据.' % (SourceFileName, Sheet.nrows-1))
    SourceFile.release_resources()
    return (Owner, Goods, Amount, Port, ArrivalDate)

### 读取分类名录中的各个子表，返回为列表，主流粉/块等返回{分类名称：品种}字典
def ReadList(ListFileName):
    GoodsClassName = {}
    GoodsClassList = {}
    ListFile = xlrd.open_workbook(ListFileName, 'r')
    ClassList = ListFile.sheets()[0].col_values(0)  #分类种类列表
    Kinds = ListFile.sheets()[1].col_values(0)  #品种列表
    SteelCompany = ListFile.sheets()[2].col_values(0)   #钢厂列表
    Trader = ListFile.sheets()[3].col_values(0)     #贸易商列表
    for i in range(3, len(ClassList)):
        GoodsClassName[i-3] = ClassList[i].replace(' ','')  #分类种类中的第4项开始为各个小的品种分类
        GoodsClassList[i-3] = ListFile.sheets()[i+1].col_values(0)
    print(u'已读取"\033[1;34;0m%s\033[0m"中的\033[1;34;0m%d\033[0m个清单.' % (ListFileName, len(ClassList)))
    ListFile.release_resources()
    return (Kinds, SteelCompany, Trader, GoodsClassName, GoodsClassList)

### 读取标准名称中的货主和品种标准名称,返回为{一般名称:标准名称}字典 ###
def ReadStd(StdFileName):
    StdOwner = {}
    StdGoods = {}
    StdFile = xlrd.open_workbook(StdFileName, 'r')
    OwnerSheet = StdFile.sheet_by_index(0)
    GoodsSheet = StdFile.sheet_by_index(1)
    for i in range(1, OwnerSheet.nrows):
        RowValue = OwnerSheet.row_values(i)
        StdOwner[RowValue[0].replace(' ','')] = RowValue[1].replace(' ','')
    print(u'已读取"\033[1;34;0m%s\033[0m"中的\033[1;34;0m%d\033[0m个货主标准名称.' % (StdFileName, OwnerSheet.nrows-1))
    for i in range(1, GoodsSheet.nrows):
        RowValue = GoodsSheet.row_values(i)
        StdGoods[RowValue[0].replace(' ','')] = RowValue[1].replace(' ','')
    print(u'已读取"\033[1;34;0m%s\033[0m"中的\033[1;34;0m%d\033[0m个品种标准名称.' % (StdFileName, GoodsSheet.nrows-1))
    StdFile.release_resources()
    return (StdOwner, StdGoods)

### 读取曹妃甸港口数据 ###
def ReadPort0(PortFilename,PortDate):
    PortFile = xlrd.open_workbook(PortFilename, formatting_info=True)

    PortFile.release_resources()

### 读取京唐港口数据 ###
def ReadPort1(PortFilename,PortDate):
    PortFile = xlrd.open_workbook(PortFilename, formatting_info=True)

    PortFile.release_resources()

### 读取岚桥港口数据 ###
def ReadPort2(PortFilename,PortDate):
    PortFile = xlrd.open_workbook(PortFilename, formatting_info=True)

    PortFile.release_resources()

### 读取岚山港口数据 ###
def ReadPort3(PortFilename,PortDate):
    PortFile = xlrd.open_workbook(PortFilename, formatting_info=True)

    PortFile.release_resources()

### 读取连云港港口数据 ###
def ReadPort4(PortFilename,PortDate):
    PortFile = xlrd.open_workbook(PortFilename, formatting_info=True)

    PortFile.release_resources()

### 读取青岛港口数据 ###
def ReadPort5(PortFilename,PortDate):
    PortFile = xlrd.open_workbook(PortFilename, formatting_info=True)

    PortFile.release_resources()

### 读取日照港口数据 ###
def ReadPort6(PortFilename,PortDate):
    PortFile = xlrd.open_workbook(PortFilename, formatting_info=True)

    PortFile.release_resources()


### 根据港口名称读取数据 ###
#'曹妃甸','京唐','岚桥','岚山','连云港','青岛','日照'
def ReadData(PortFilename,PortShortname,PortDate):
    PortData = []
    if PortShortname == '曹妃甸': PortData = ReadPort0(PortFilename,PortDate)
    elif PortShortname == '京唐': PortData = ReadPort1(PortFilename,PortDate)
    elif PortShortname == '岚桥': PortData = ReadPort2(PortFilename,PortDate)
    elif PortShortname == '岚山': PortData = ReadPort3(PortFilename,PortDate)
    elif PortShortname == '连云港': PortData = ReadPort4(PortFilename,PortDate)
    elif PortShortname == '青岛': PortData = ReadPort5(PortFilename,PortDate)
    elif PortShortname == '日照': PortData = ReadPort6(PortFilename,PortDate)
    else: print(PortFilename,'港口名称或港口名称不在标准名称中')
    return PortData

### 读取各港口数据 ###
def CollectPortData(PortFilesList,StdPort):
    AllData = []    #最终存储的数据，[货主，品种，库存，港口，到港日期]
    PortDates = GetFileDate(PortFilesList, StdPort)
    PortShortnames = GetShortName(PortFilesList,StdPort)
    for i in range(0,len(PortFilesList)):
        PortFilename = PortFilesList[i]
        PortData = ReadData(PortFilename,PortShortnames[i],PortDates[i])
        print(PortFilename)
        print(PortData)
        #AllData.append(PortData)