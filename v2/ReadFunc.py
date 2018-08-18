# coding=utf-8
import xlrd

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
