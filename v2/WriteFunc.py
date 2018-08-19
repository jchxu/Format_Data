# coding=utf-8
import xlsxwriter

def CalcRatio(total, steel, trade):
    steelratio = 0
    traderatio = 0
    if total != 0:
        steelratio = 100 * steel / total
        traderatio = 100 * trade / total
    return (steelratio, traderatio)

### 打印输出各港口、分类、品种汇总数量、货权集中度 （屏幕输出、csv、txt版）###
def WriteSummary(AmountInfo,ShipInfo, PortList, GoodsClassName, GoodsClassList):
#AmountInfo = [TotalAmount, TotalSteel, ClassTotal, ClassSteel, GoodsTotal, GoodsSteel]
#ShipInfo = [GoodShip, GoodSteelShip, GoodOtherShip]
    for i in range(0,len(PortList)):
        print(PortList[i])
        port = PortList[i]
        for j in range(0,len(GoodsClassName)):
            classname = GoodsClassName[j]
            calsstotal = AmountInfo[2][port][classname]
            classsteel = AmountInfo[3][port][classname]
            classtrade = calsstotal - classsteel
            steelratio, traderatio = CalcRatio(calsstotal, classsteel, classtrade)
            print("%s,%.0f,%.0f,%.1f%%,%.0f,%.1f%%" % (classname,calsstotal,classsteel, steelratio, classtrade, traderatio), sep=',')
            if (j == 0) or (j == 1):
                for k in range(0,len(GoodsClassList[j])):
                    goodname = GoodsClassList[j][k]
                    if goodname in AmountInfo[4][port].keys():
                        goodtotal = AmountInfo[4][port][goodname]
                        goodsteel = AmountInfo[5][port][goodname]
                        goodtrade = goodtotal - goodsteel
                    else:
                        goodtotal = 0
                        goodsteel = 0
                        goodtrade = 0
                    steelratio, traderatio = CalcRatio(goodtotal, goodsteel, goodtrade)
                    print("%s,%.0f,%.0f,%.1f%%,%.0f,%.1f%%" % (goodname, goodtotal, goodsteel, steelratio, goodtrade, traderatio), sep=',')

