# coding=utf-8
import re

### 将货主/品种名称标准化，若不在标准化名称中，则输出提示 ###
def Standardize(Owner, Goods, StdOwner, StdGoods):
    NoStdOwner = []
    NoStdGoods = []
    Flag = 0
    for i in range(0,len(Owner)):
        if Owner[i] in StdOwner.keys():
            Owner[i] = StdOwner[Owner[i]]
        elif (not (Owner[i] in StdOwner.keys())) and (not (Owner[i] in StdOwner.values())):
            NoStdOwner.append(Owner[i])
    if NoStdOwner:
        Flag = 1
        NoStdOwner = list(set(NoStdOwner))
        print(u'\033[1;34;0m%d\033[0m个货主名称不在标准名称中: %r' % (len(NoStdOwner),NoStdOwner))
    for i in range(0, len(Goods)):
        if Goods[i] in StdGoods.keys():
            Goods[i] = StdGoods[Goods[i]]
        elif (not (Goods[i] in StdGoods.keys())) and (not (Goods[i] in StdGoods.values())):
            NoStdGoods.append(Goods[i])
    if NoStdGoods:
        Flag = 1
        NoStdGoods = list(set(NoStdGoods))
        print(u'\033[1;34;0m%d\033[0m个品种名称不在标准名称中: %r' % (len(NoStdGoods), NoStdGoods))
    if Flag == 1:
        print(u'\033[1;34;0m请首先更新标准名称清单，程序退出!\033[0m')
        #exit()
    else:
        print(u'已完成货主/品种名称标准化')
    return (Owner, Goods)

### 根据港口标准名称分类数据，返回为字典 ###
def ClassifyByPort(Owner, Goods, Amount, Port, StdPort):
    OwnerByPort = {}
    GoodsByPort = {}
    AmountByPort = {}
    for item in StdPort:
        OwnerByPort[item] = []
        GoodsByPort[item] = []
        AmountByPort[item] = []
    for i in range(0,len(Owner)):
        if Port[i] in StdPort:
            OwnerByPort[Port[i]].append(Owner[i])
            GoodsByPort[Port[i]].append(Goods[i])
            AmountByPort[Port[i]].append(Amount[i])
    return (OwnerByPort,GoodsByPort,AmountByPort)
