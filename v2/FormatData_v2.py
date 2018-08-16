# coding=utf-8
from ReadFunc import *
from OperateFunc import *
from WriteFunc import *

### 需要用户定义的参数 ###
SourceFile = "京唐港01.07库存.xlsx"   #港口数据文件
ListFile = "分类名录.xlsx"  #记录主流粉矿、主流块矿、非主流资源、品种、钢厂、贸易商名录的文件
StdFile = "标准名称.xlsx"  #记录货主（钢厂、贸易商）、品种标准名称的数据文件

### 初始化，读取原始数据 ###
Owner, Goods, Amount, Port, ArrivalDate = ReadSource(SourceFile)    #读取港口库存数据
Kinds, SteelCompany, Trader, GoodsClassName, GoodsClassList = ReadList(ListFile)    #读取分类名录中的各个子表，返回为列表，主流粉/块等返回{分类名称：品种}字典
StdOwner, StdGoods = ReadStd(StdFile)   #读取标准名称中的货主和品种标准名称

### 数据处理 ###
Owner, Goods = Standardize(Owner, Goods, StdOwner, StdGoods)    #货主/品种名称标准化
