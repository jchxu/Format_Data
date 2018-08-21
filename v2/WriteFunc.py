# coding=utf-8
import xlsxwriter

### 计算百分比，返回*100的结果，总数为0则全部返回0 ###
def CalcRatio(Total, Steel, Trade):
    SteelRatio = 0
    TradeRatio = 0
    if Total != 0:
        SteelRatio = 100 * Steel / Total
        TradeRatio = 100 * Trade / Total
    return (SteelRatio, TradeRatio)

### TopShip格式化： 货主(数量)，货主(数量)……###
def TopShipFormat(TopList):
    SepStr = '/'
    List = []
    for i in range(0,len(TopList)):
        List.append('%s(%.0f)' % (TopList[i][0],TopList[i][1]))
    return (SepStr.join(List))

### 统计货权排名（按数量统计排名） ###
def TopShip(GoodShipDict):
    GoodTop13 = []
    GoodTop46 = []
    GoodTopOther = []
    Sortedlist = sorted(GoodShipDict.items(),key = lambda x:x[1],reverse = True)
    if len(Sortedlist) <= 3:
        GoodTop13 = Sortedlist
    elif 3 < len(Sortedlist) <= 6:
        GoodTop13 = Sortedlist[0:3]
        GoodTop46 = Sortedlist[3:6]
    elif len(Sortedlist) > 6:
        GoodTop13 = Sortedlist[0:3]
        GoodTop46 = Sortedlist[3:6]
        GoodTopOther = Sortedlist[6:len(Sortedlist)]
    GoodTop13 = TopShipFormat(GoodTop13)
    GoodTop46 = TopShipFormat(GoodTop46)
    GoodTopOther = TopShipFormat(GoodTopOther)
    return (GoodTop13,GoodTop46,GoodTopOther)

### 打印输出各港口、分类、品种汇总数量、货权集中度 （屏幕输出、csv、txt版）###
#AmountInfo = [TotalAmount, TotalSteel, ClassTotal, ClassSteel, GoodsTotal, GoodsSteel]
#ShipInfo = [GoodShip, GoodSteelShip, GoodOtherShip]
#def WriteSummary(AmountInfo,ShipInfo, PortList, GoodsClassName, GoodsClassList):
#    for i in range(0,len(PortList)):
#        print(PortList[i])
#        port = PortList[i]
#        totalamount = AmountInfo[0][port]
#        totalsteel = AmountInfo[1][port]
#        totaltrade = totalamount - totalsteel
#        steeltotalratio, tradetotalratio = CalcRatio(totalamount, totalsteel, totaltrade)
#        print("%s,%.0f,%.0f,%.1f%%,%.0f,%.1f%%" % ('合计', totalamount, totalsteel, steeltotalratio, totaltrade, tradetotalratio), sep=',')
#        for j in range(0,len(GoodsClassName)):
#            classname = GoodsClassName[j]
#            calsstotal = AmountInfo[2][port][classname]
#            classsteel = AmountInfo[3][port][classname]
#            classtrade = calsstotal - classsteel
#            steelratio, traderatio = CalcRatio(calsstotal, classsteel, classtrade)
#            print("%s,%.0f,%.0f,%.1f%%,%.0f,%.1f%%" % (classname,calsstotal,classsteel, steelratio, classtrade, traderatio), sep=',')
#            if (j == 0) or (j == 1): #pass
#                for k in range(0,len(GoodsClassList[j])):
#                    goodname = GoodsClassList[j][k]
#                    if goodname in AmountInfo[4][port].keys():
#                        goodtotal = AmountInfo[4][port][goodname]
#                        goodsteel = AmountInfo[5][port][goodname]
#                        goodtrade = goodtotal - goodsteel
#                        goodship = ShipInfo[0][port][goodname]
#                        goodsteelship = ShipInfo[1][port][goodname]
#                        goodothership = ShipInfo[2][port][goodname]
#                        GoodTop13, GoodTop46, GoodTopOther = TopShip(goodship)
#                        GoodSteelTop13, GoodSteelTop46, GoodSteelTopOther = TopShip(goodsteelship)
#                        GoodOtherTop13, GoodOtherTop46, GoodOtherTopOther = TopShip(goodothership)
#                    else:
#                        goodtotal = 0
#                        goodsteel = 0
#                        goodtrade = 0
#                        GoodTop13, GoodTop46, GoodTopOther = ('','','')
#                    steelratio, traderatio = CalcRatio(goodtotal, goodsteel, goodtrade)
#                    print("%s,%.0f,%.0f,%.1f%%,%.0f,%.1f%%" % (goodname, goodtotal, goodsteel, steelratio, goodtrade, traderatio), sep=',', end = ',')
#                    print("%s,%s,%s,%s,%s,%s,%s,%s,%s" % (GoodTop13, GoodTop46, GoodTopOther,GoodSteelTop13, GoodSteelTop46, GoodSteelTopOther,GoodOtherTop13, GoodOtherTop46, GoodOtherTopOther), sep=',')
#
###
def WriteSummary(AmountInfo,ShipInfo, PortList, GoodsClassName, GoodsClassList):
    Filename = 'test.xlsx'
    WorkBook = xlsxwriter.Workbook(Filename)
    Sheet1 = WorkBook.add_worksheet()

    for i in range(0, len(PortList)):
    # print(PortList[i])
        port = PortList[i]
        Sheet1.write(0,i*9,port)

    WorkBook.close()
