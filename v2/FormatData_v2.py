# coding=utf-8
from ReadFunc import *
from OperateFunc import *
from WriteFunc import *

### 需要用户定义的参数 ###
SourceFile = "京唐港8.1库存.xlsx"   #港口数据文件
SourceFile = "京唐港01.07库存.xlsx"   #港口数据文件
ListFile = "分类名录.xlsx"  #记录主流粉矿、主流块矿、非主流资源、品种、钢厂、贸易商名录的文件
StdFile = "标准名称.xlsx"  #记录货主（钢厂、贸易商）、品种标准名称的数据文件
StdPort = [u'曹妃甸',u'京唐',u'岚桥',u'岚山',u'连云港',u'青岛',u'日照']

### 初始化，读取原始数据 ###
Owner, Goods, Amount, Port, ArrivalDate = ReadSource(SourceFile)    #读取港口库存数据
Kinds, SteelCompany, Trader, GoodsClassName, GoodsClassList = ReadList(ListFile)    #读取分类名录中的各个子表，返回为列表，主流粉/块等返回{分类名称：品种}字典
StdOwner, StdGoods = ReadStd(StdFile)   #读取标准名称中的货主和品种标准名称
print(GoodsClassName, GoodsClassList)
### 数据处理 ###
Owner, Goods, Port = Standardize(Owner, Goods, Port, StdOwner, StdGoods, StdPort)    #货主/品种/港口名称标准化
PortList, GoodList, DataClassified = ClassifyByPortKind(Owner, Goods, Amount, Port)  #根据港口和品种分类数据
AmountInfo = SummaryAmount(PortList, GoodList, DataClassified, SteelCompany, GoodsClassName, GoodsClassList)
ShipInfo = SummaryShip(PortList, GoodList, DataClassified, SteelCompany, GoodsClassName, GoodsClassList)

### 数据输出 ###
WriteSummary('csv',AmountInfo,ShipInfo, PortList, GoodsClassName, GoodsClassList)