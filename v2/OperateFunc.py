# coding=utf-8
#import re

### 将货主/品种/港口名称标准化，若不在标准化名称中，则输出提示 ###
def Standardize(Owner, Goods, Port, StdOwner, StdGoods, StdPort):
    NoStdOwner = []
    NoStdGoods = []
    NoStdPort = []
    Flag = 0
    #货主名称标准化
    for i in range(0,len(Owner)):
        if Owner[i] in StdOwner.keys():
            Owner[i] = StdOwner[Owner[i]]
        elif (not (Owner[i] in StdOwner.keys())) and (not (Owner[i] in StdOwner.values())):
            NoStdOwner.append(Owner[i])
    if NoStdOwner:  #输出去重的非标准化货主名称
        Flag = 1
        NoStdOwner = list(set(NoStdOwner))
        print(u'\033[1;34;0m%d\033[0m个货主名称不在标准名称中: %r' % (len(NoStdOwner),NoStdOwner))
    #品种名称标准化
    for i in range(0, len(Goods)):
        if Goods[i] in StdGoods.keys():
            Goods[i] = StdGoods[Goods[i]]
        elif (not (Goods[i] in StdGoods.keys())) and (not (Goods[i] in StdGoods.values())):
            NoStdGoods.append(Goods[i])
    if NoStdGoods:  #输出去重的非标准化品种名称
        Flag = 1
        NoStdGoods = list(set(NoStdGoods))
        print(u'\033[1;34;0m%d\033[0m个品种名称不在标准名称中: %r' % (len(NoStdGoods), NoStdGoods))
    #港口名称标准化
    for i in range(0,len(Port)):
        for item in StdPort:
            if item in Port[i]:
                Port[i] = item
    #提示更新标准化名称
    if Flag == 1:
        print(u'\033[1;34;0m请首先更新标准名称清单，程序退出!\033[0m')
        #exit()
    else:
        print(u'已完成货主/品种名称标准化')
    return (Owner, Goods, Port)

### 根据港口标准名称分类数据，返回为多维字典 {港口:{品种:{货主:数量}}} ###
def ClassifyByPortKind(Owner, Goods, Amount, Port):
    DataClassified = {}
    GoodList = list(set(Goods))
    PortList = sorted(list(set(Port)))
    #初始化分类数据字典
    for i in PortList:
        DataClassified[i] = {}
        for j in GoodList:
            DataClassified[i][j] = {}
    # 根据港口/品种/货主分类数据
    for i in range(0,len(Owner)):
        DataClassified[Port[i]][Goods[i]][str(i)+'-'+Owner[i]] = Amount[i]
    return (PortList, GoodList, DataClassified)

### 根据{货主:数量}字典，统计总数量、钢厂总数量、贸易商总数量
def SumTotal(GoodDataDict, SteelCompany):
    Total = 0
    Steel = 0
    for item in GoodDataDict.keys():
        Total += GoodDataDict[item]
        if item.split('-')[1] in SteelCompany:
            Steel += GoodDataDict[item]
    return (Total, Steel)

### 根据港口、分类、品种汇总数量 ###
def SummaryAmount(PortList, GoodList, DataClassified, SteelCompany, GoodsClassName, GoodsClassList):
    TotalAmount = {}   #{港口:总数量}
    TotalSteel = {}    #{港口:港口总数量}
    ClassTotal = {}    #{港口:{分类名称:总数量}}
    ClassSteel = {}    #{港口:{分类名称:钢厂总数量}}
    GoodsTotal = {}    #{港口:{品种名称:总数量}}
    GoodsSteel = {}    #{港口:{品种名称:钢厂总数量}}
    for i in PortList:
        ClassTotal[i] = {}
        ClassSteel[i] = {}
        GoodsTotal[i] = {}
        GoodsSteel[i] = {}
    #按港口/品种统计
    for i in PortList:
        for j in DataClassified[i].keys():
            GoodData = DataClassified[i][j]
            if len(GoodData) == 0:  #该港口、该品种没有数据
                GoodsTotal[i][j] = 0
                GoodsSteel[i][j] = 0
            else:    #该港口、该品种有一个或多个货主数据
                GoodsTotal[i][j], GoodsSteel[i][j] = SumTotal(GoodData, SteelCompany)
            #print(i,j,GoodsTotal[i][j],GoodsSteel[i][j])
        #分类统计汇总
        for j in GoodsClassName.keys():
            ClassTotal[i][GoodsClassName[j]] = 0
            ClassSteel[i][GoodsClassName[j]] = 0
        for j in GoodsClassName.keys():
            #print(i,GoodsClassName[j],GoodsClassList[j])
            for k in GoodsClassList[j]:
                if k in GoodsTotal[i].keys():
                    ClassTotal[i][GoodsClassName[j]] += GoodsTotal[i][k]
                    ClassSteel[i][GoodsClassName[j]] += GoodsSteel[i][k]
            #print(i,GoodsClassName[j],ClassTotal[i][GoodsClassName[j]],ClassSteel[i][GoodsClassName[j]])
        TotalAmount[i] = sum(list(GoodsTotal[i].values()))
        TotalSteel[i] = sum(list(GoodsSteel[i].values()))
    #按分类统计汇总
    #print(GoodsClassName)
    #print(GoodsClassList)
    AmountInfo = [TotalAmount, TotalSteel, ClassTotal, ClassSteel, GoodsTotal, GoodsSteel]
    return (AmountInfo)

### 根据{品种：{货主:数量}}字典，统计不同货主的品种数量 ###
def CalcShip(GoodDataDict, SteelCompany):
    GoodOwner = []
    GoodShip = {}
    GoodSteelShip = {}
    GoodOtherShip = {}
    for item in GoodDataDict.keys():
        GoodOwner.append(item.split('-')[1])
    GoodOwner = list(set(GoodOwner))    #非重复的货主列表
    for item in GoodOwner:
        GoodShip[item] = 0
        if item in SteelCompany:
            GoodSteelShip[item] = 0
        else:
            GoodOtherShip[item] = 0
    for item in GoodDataDict.keys():
        Owner = item.split('-')[1]
        GoodShip[Owner] += GoodDataDict[item]
        if Owner in SteelCompany:
            GoodSteelShip[Owner] += GoodDataDict[item]
        else:
            GoodOtherShip[Owner] += GoodDataDict[item]
    return (GoodShip,GoodSteelShip,GoodOtherShip)

### 统计各港口、品种下，货主/钢厂/贸易商的数量 ###
def SummaryShip(PortList, GoodList, DataClassified, SteelCompany, GoodsClassName, GoodsClassList):
    ClassShip = {}      #{港口:{分类:{货主：数量}}}
    ClassSteelShip = {} #{港口:{分类:{钢厂：数量}}}
    ClassOtherShip = {} #{港口:{分类:{贸易商：数量}}}
    GoodShip = {}       #{港口:{品种:{货主：数量}}}
    GoodSteelShip = {}  #{港口:{品种:{钢厂：数量}}}
    GoodOtherShip = {}  #{港口:{品种:{贸易商：数量}}}
    for i in PortList:
        GoodShip[i] = {}
        GoodSteelShip[i] = {}
        GoodOtherShip[i] = {}
    for i in PortList:
        for j in DataClassified[i].keys():
            GoodShip[i][j] = {}
            GoodSteelShip[i][j] = {}
            GoodOtherShip[i][j] = {}
        for j in DataClassified[i].keys():
            GoodData = DataClassified[i][j]
            GoodShip[i][j], GoodSteelShip[i][j], GoodOtherShip[i][j] = CalcShip(GoodData, SteelCompany)
    ShipInfo = [GoodShip, GoodSteelShip, GoodOtherShip]
    return (ShipInfo)