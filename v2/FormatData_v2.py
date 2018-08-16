# coding=utf-8
from ReadFunc import *

### 需要用户定义的变量 ###
SourceFile = "京唐港01.07库存.xlsx"   # Excel数据文件的文件名，带扩展名。
ListFile = "分类名录.xlsx"  # 记录主流粉矿、主流块矿、非主流资源、品种、钢厂、贸易商名录的文件
StdFile = "标准名称.xlsx"  # 记录货主（钢厂、贸易商）、品种标准名称的数据文件

### 初始化 ###
Kinds, SteelCompany, Trader, GoodsClassName, GoodsClassList = ReadList(ListFile)    # 读取分类名录中的各个子表，返回为列表，主流粉/块等返回{分类名称：品种}字典
