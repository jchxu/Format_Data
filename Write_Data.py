# coding=utf-8
import xlrd, xlwt, re
from Read_Data import *

"""
transfer_nick_amount
    v1.0: 将数量数据中的“万”“吨”等转换为统一格式
standardize_name
    v1.0: 更改货主、货物中的昵称为标准名称
sum_owner_goods
    v1.0: 返回指定货主、指定货物的总数量
getratio
    v1.0: 判断numa是否为0，然后计算numa/numb，若为0则返回/
getgoodsamount
    v1.0: 返回指定货物的总量、钢厂总量、钢厂占比、贸易商总量、贸易商占比
    v1.1: 更改钢厂、贸易商判断方式：若非钢厂，则为贸易商
calculate_summary
    v1.0: 计算统计数据结果
write_summary
    v1.0: 以summary形式输出统计结果
write_detail
    v1.0: 依次写入序号、货主、货种、数量的数据，每条数据一行
calculate_trackdata
    v1.0: 统计历史追踪数据中的各粉矿、块矿数据
write_tracking
    v1.0: 输出历史追踪数据
"""

def transfer_nick_amount(subsheet, colindex):
    titleindex = 0
    ownerindex = 0
    goodsindex = 0
    amountindex = 0
    ### 确定标题行位置 ###
    for j in range(0, subsheet.nrows):
        line = subsheet.row_values(j)
        if u"船名" in line or u"船舶名称" in line or u"货主" in line:
            titleindex = j
            break
    title = subsheet.row_values(titleindex)
    ### 确定货物、货主、数量列位置 ###
    for k in range(0, len(title)):
        if title[k] in [u"货名", u"货种", u"货性", u"品种"]:
            goodsindex = k
        elif title[k] in [u"货主"] or (u"收货人" in title[k]) or (u"钢厂" in title[k]):
            ownerindex = k
        elif title[k] in [u"结 存 量", u"库存", u"港存数"] or u"数量" in title[k]:
            amountindex = k
    #print ownerindex, goodsindex, amountindex
    ### 遍历每行，检查是否是所需数据，是则读取入相应字典 ###
    amount = {}
    for k in range(titleindex+1, subsheet.nrows):
        data = subsheet.row_values(k)
        if data[goodsindex] and (not (u"合计" in data[goodsindex])) \
                and isinstance(data[amountindex], float):  # 判断货物和数量列是否有数据，且不是合计的数据。
            cell = data[colindex-1]
            if cell.count('#') <= 1:
                amountnick = cell.split('#')[-1]
                number = float(re.findall(r"\d+\.?\d*", amountnick)[0])
                if (u"万" in amountnick) or (u"万吨" in amountnick):
                    amount[k] = number * 10000
                elif (u"千" in amountnick) or (u"千吨" in amountnick):
                    amount[k] = number * 1000
                elif u"吨" in amountnick:
                    amount[k] = number
                elif (type(eval(amountnick)) == float) or (type(eval(amountnick)) == int):
                    amount[k] = number
            elif cell.count('#') > 1:
                number = 0
                for item in cell.split('#'):
                    if (u"万" in item) or (u"万吨" in item):
                        number += float(re.findall(r"\d+\.?\d*",item)[0]) * 10000
                    elif (u"千" in item) or (u"千吨" in item):
                        number += float(re.findall(r"\d+\.?\d*",item)[0]) * 1000
                    elif u"吨" in item:
                        number += float(re.findall(r"\d+\.?\d*",item)[0])
                amount[k] = number
    return amount

def standardize_name(stdname, owner, goods):
    stdfile = xlrd.open_workbook(stdname.decode('utf-8'), 'r')
    ownerdic = {}
    goodsdic = {}
    ownersheet = stdfile.sheet_by_name("owner")
    goodssheet = stdfile.sheet_by_name("goods")
    ### 读入标准化名称，以字典形式保存 ###
    for i in range(1, ownersheet.nrows):
        rowValue = ownersheet.row_values(i)
        ownerdic[rowValue[0]] = rowValue[1]
    for i in range(1, goodssheet.nrows):
        rowValue = goodssheet.row_values(i)
        goodsdic[rowValue[0]] = rowValue[1]
    ### 更改货主、货物中的昵称为标准名称 ###
    for item in owner.keys():
        if owner[item] in ownerdic.keys():
            owner[item] = ownerdic[owner[item]]
    for item in goods.keys():
        if goods[item] in goodsdic.keys():
            goods[item] = goodsdic[goods[item]]
    print "All names of owners and goods have been standardized."
    return (owner, goods)

def sum_owner_goods (owner, goods, amount, ownername, goodsname):
    """依次输入存储收货人名称、货物名称、数量的字典，指定的货主、货物（所有用ALL表示），
    遍历求和，返回指定货主、指定货物的总数量"""
    summation = 0
    if ownername == "ALL" and goodsname == "ALL":   #所有货主的所有货物
        for i in amount.values():
            summation += i
    elif ownername == "ALL" and goodsname != "ALL":   #所有货主的指定货物
        for i in goods.keys():
            if goods[i] == goodsname:
                summation += amount[i]
    elif ownername != "ALL" and goodsname == "ALL":   # 指定货主的所有货物
        for i in owner.keys():
            if owner[i] == ownername:
                summation += amount[i]
    else:                                             # 指定货主的指定货物
        for i in owner.keys():
            if owner[i] == ownername and goods[i] == goodsname:
                summation += amount[i]
    return summation

def getratio(numa, numb):
    """判断numa是否为0，然后计算numa / numb"""
    if numa != 0:
        return numa/numb
    else:
        return "/"

def getgoodsamount(goodsname, owner, goods, amount, company):
    """返回指定货物的总量、钢厂总量、钢厂占比、贸易商总量、贸易商占比"""
    totalamount = sum_owner_goods(owner,goods,amount, u"ALL", goodsname)
    companyamount = 0
    traderamount = 0
    for item in set(owner.values()):
        if item in company:
            companyamount += sum_owner_goods(owner,goods,amount, item, goodsname)
        else:
            traderamount += sum_owner_goods(owner,goods,amount, item, goodsname)
    companyratio = getratio(companyamount, totalamount)
    traderratio = getratio(traderamount, totalamount)
    return (totalamount, companyamount, companyratio, traderamount, traderratio)

def calculate_summary(mainpowder, mainblock, nonmain, owner, goods, amount, company):
    ### 计算所有货物的总和 ###
    totalamount = sum_owner_goods(owner, goods, amount, u"ALL", u"ALL")
    totalcom = 0
    totaltrader = 0
    for item in set(owner.values()):
        if item in company:
            totalcom += sum_owner_goods(owner, goods, amount, item, u"ALL")
        else:
            totaltrader += sum_owner_goods(owner, goods, amount, item, u"ALL")
    if (totalamount-totalcom-totaltrader) > 1:
        print "The Lists of Companies and Traders May Be IN-Complete!"
        for i in range(1, len(owner.keys())+1):
            if (not (owner[i] in company)) and (not (owner[i] in trader)):
                print 'No.%d: The "%s" Was Not In the List of Company or Trader!' % (i, owner[i])
    elif (totalamount-totalcom-totaltrader) < -1:
        for i in range(1, len(owner.keys())+1):
            if (owner[i] in company) and (owner[i] in trader):
                print 'No.%d: The "%s" Was In the List of Company, Meanwhile, In List of Trader!' % (i, owner[i])
    totalcomratio = getratio(totalcom, totalamount)
    totaltraderratio = getratio(totaltrader, totalamount)
    #print totalamount, totalcom, totalcomratio, totaltrader, totaltraderratio
    totalrow = [u"港存总计", totalamount, totalcom, totalcomratio, totaltrader, totaltraderratio]

    ### 计算每种需单独统计的货物的总和 ###
    goodstotal = {}
    goodscom = {}
    goodscomratio = {}
    goodstrader = {}
    goodstraderratio = {}
    goodsrow = {}
    biglist = mainpowder + mainblock + nonmain
    for i in range(0, len(biglist)):
        out = getgoodsamount(biglist[i], owner, goods, amount, company)
        goodstotal[i], goodscom[i], goodscomratio[i], goodstrader[i], goodstraderratio[i] = out
        #print goodstotal[i], goodscom[i], goodscomratio[i], goodstrader[i], goodstraderratio[i]
        goodsrow[i] = [biglist[i], goodstotal[i], goodscom[i], goodscomratio[i], goodstrader[i], goodstraderratio[i]]

    ### 计算主流粉矿的总和 ###
    powdertotal =0
    powdercom =0
    powdertrader = 0
    for i in range(0, len(mainpowder)):
        powdertotal += goodstotal[i]
        powdercom += goodscom[i]
        powdertrader += goodstrader[i]
    powdercomratio = getratio(powdercom, powdertotal)
    powdertraderratio = getratio(powdertrader, powdertotal)
    #print powdertotal, powdercom, powdercomratio, powdertrader, powdertraderratio
    powderrow = [u"主流粉矿", powdertotal, powdercom, powdercomratio, powdertrader, powdertraderratio]

    ### 计算主流块矿的总和 ###
    blocktotal = 0
    blockcom = 0
    blocktrader = 0
    for i in range(0, len(mainblock)):
        blocktotal += goodstotal[i+len(mainpowder)]
        blockcom += goodscom[i+len(mainpowder)]
        blocktrader += goodstrader[i+len(mainpowder)]
    blockcomratio = getratio(blockcom, blocktotal)
    blocktraderratio = getratio(blocktrader, blocktotal)
    #print blocktotal, blockcom, blockcomratio, blocktrader, blocktraderratio
    blockrow = [u"主流块矿", blocktotal, blockcom, blockcomratio, blocktrader, blocktraderratio]

    ### 计算主流资源（主流粉矿+主流块矿） ###
    maintotal = powdertotal + blocktotal
    maincom = powdercom + blockcom
    maintrader = powdertrader + blocktrader
    maincomratio = getratio(maincom, maintotal)
    maintraderratio = getratio(maintrader, maintotal)
    #print maintotal, maincom, maincomratio, maintrader, maintraderratio
    mainrow = [u"主流资源", maintotal, maincom, maincomratio, maintrader, maintraderratio]

    ### 计算非主流资源（总的-主流-乌克兰精粉-乌克兰球团） ###
    nonmaintotal = totalamount - maintotal - \
                   goodstotal[len(mainpowder)+len(mainblock)] - goodstotal[len(mainpowder)+len(mainblock)+1]
    nonmaincom = totalcom - maincom - \
                   goodscom[len(mainpowder)+len(mainblock)] - goodscom[len(mainpowder)+len(mainblock)+1]
    nonmaintrader = totaltrader - maintrader - \
                 goodstrader[len(mainpowder) + len(mainblock)] - goodstrader[len(mainpowder) + len(mainblock) + 1]
    nonmaincomratio = getratio(nonmaincom, nonmaintotal)
    nonmaintraderratio = getratio(nonmaintrader, nonmaintotal)
    #print nonmaintotal, nonmaincom, nonmaincomratio, nonmaintrader, nonmaintraderratio
    nonmainrow = [u"非主流资源", nonmaintotal, nonmaincom, nonmaincomratio, nonmaintrader, nonmaintraderratio]

    return (totalrow, mainrow, nonmainrow, powderrow, blockrow, goodsrow)

def write_summary(resultfile, mainpowder, mainblock, nonmain, totalrow, mainrow, nonmainrow, powderrow, blockrow, goodsrow):
    subsheet = resultfile.add_sheet("Summary")
    titlerow = [u"矿种", u"合计", u"钢厂", u"钢厂占比", u"贸易商", u"贸易商占比"]
    powderadd = 5   #从第5行开始写主流粉矿,即主流资源与主流粉矿之间空1行
    blockadd = 7 + len(mainpowder)  # 加2表示空1行开始写主流块矿。+n表示空n-1行
    nonmainadd = 9 + len(mainpowder) + len(mainblock)   # 再加2（n）表示空1（n-1）行开始写非主流资源。再+n表示再空n-1行

    for i in range(0, len(titlerow)):
        subsheet.write(0, i, titlerow[i])
        subsheet.write(1, i, totalrow[i])
        subsheet.write(2, i, mainrow[i])
        subsheet.write(powderadd-1, i, powderrow[i])
        for k in range(0, len(mainpowder)):
            subsheet.write(powderadd+k, i, goodsrow[k][i])
        subsheet.write(blockadd-1, i, blockrow[i])
        for k in range(0, len(mainblock)):
            subsheet.write(blockadd+k, i, goodsrow[k+len(mainpowder)][i])
        subsheet.write(nonmainadd-1, i, nonmainrow[i])
        for k in range(0, len(nonmain)):
            subsheet.write(nonmainadd+k, i, goodsrow[k+len(mainpowder)+len(mainblock)][i])
    print 'Summary Data Have Been Written in Subsheet "%s".' % subsheet.name.encode('utf-8')
    return resultfile

def write_detail(resultfile, owner, goods, amount):
    """按照序号、货主、货物、数量的顺序写入所有数据"""
    subsheet = resultfile.add_sheet("Detail")
    titlerow = [u"序号", u"货主", u"货名", u"数量"]

    ### 写标题行 ###
    for i in range(0, len(titlerow)):
        subsheet.write(0, i, titlerow[i])
    ### 写每行数据 ###
    for i in range(1, len(owner.keys()) + 1):
        subsheet.write(i, 0, i)
        subsheet.write(i, 1, owner[i])
        subsheet.write(i, 2, goods[i])
        subsheet.write(i, 3, amount[i])
    print 'Detail data have been written in subsheet "%s".' % subsheet.name.encode('utf-8')
    return resultfile

def calculate_trackdata(powder, block, goodsrow, owner, goods, amount, company):
    goodslist = powder + block
    goodsname = {}
    goodsindex = {}
    goodsdata = {}
    for i in range(0,len(goodsrow)):
        goodsname[i] = goodsrow[i][0]
        goodsindex[goodsrow[i][0]] = i
    for i in range(0, len(goodslist)):
        if goodslist[i] in goodsname.values():
            k = goodsindex[goodslist[i]]
            goodsdata[i] = [goodsrow[k][1], goodsrow[k][2], goodsrow[k][3], goodsrow[k][4],goodsrow[k][5]]
        else:
            goodsdata[i] = list(getgoodsamount(goodslist[i], owner, goods, amount, company))
    return goodsdata

def write_tracking(stddate, trackfile, subsheet, rowindex, powder, block, totalrow, mainrow, nonmainrow, powderrow, blockrow, goodsdata):
    """追加输出历史追踪数据"""
    year = get_date_time()[1]
    month = stddate[0:2]
    day = stddate[2:4]
    date = "%4s/%2s/%2s" % (year, month, day)
    # 标题部分
    if rowindex == 0:
        rowindex += 4
        subsheet.write(rowindex, 0, u"日期")
        subsheet.write(rowindex, 1, u"总库存")
        subsheet.write(rowindex, 2, u"主流")
        subsheet.write(rowindex, 3, u"主流占比")
        subsheet.write(rowindex, 4, u"粉矿")
        subsheet.write(rowindex, 5, u"粉矿占比")
        for i in range(0, len(powder)):
            subsheet.write(rowindex, 6 + i * 6, powder[i])
            subsheet.write(rowindex, 7 + i * 6, powder[i]+u"占比")
            subsheet.write(rowindex, 8 + i * 6, u"钢厂"+powder[i])
            subsheet.write(rowindex, 9 + i * 6, u"钢厂"+powder[i]+u"占比")
            subsheet.write(rowindex, 10 + i * 6, u"贸易商"+powder[i])
            subsheet.write(rowindex, 11 + i * 6, u"贸易商"+powder[i]+u"占比")
        subsheet.write(rowindex, 6+6*len(powder), u"块矿")
        subsheet.write(rowindex, 7+6*len(powder), u"块矿占比")
        for i in range(0, len(block)):
            subsheet.write(rowindex, 8+6*len(powder) + i * 6, block[i])
            subsheet.write(rowindex, 9+6*len(powder) + i * 6, block[i]+u"占比")
            subsheet.write(rowindex, 10+6*len(powder) + i * 6, u"钢厂"+block[i])
            subsheet.write(rowindex, 11+6*len(powder) + i * 6, u"钢厂"+block[i]+u"占比")
            subsheet.write(rowindex, 12+6*len(powder) + i * 6, u"贸易商"+block[i])
            subsheet.write(rowindex, 13+6*len(powder) + i * 6, u"贸易商"+block[i]+u"占比")
        subsheet.write(rowindex, 8 + 6 * (len(powder)+len(block)), u"钢厂资源")
        subsheet.write(rowindex, 9 + 6 * (len(powder)+len(block)), u"钢厂占比")
        subsheet.write(rowindex, 10 + 6 * (len(powder)+len(block)), u"贸易商资源")
        subsheet.write(rowindex, 11 + 6 * (len(powder)+len(block)), u"贸易商占比")
        subsheet.write(rowindex, 12 + 6 * (len(powder) + len(block)), u"非主流")
        subsheet.write(rowindex, 13 + 6 * (len(powder) + len(block)), u"钢厂资源")
        subsheet.write(rowindex, 14 + 6 * (len(powder) + len(block)), u"钢厂占比")
        subsheet.write(rowindex, 15 + 6 * (len(powder) + len(block)), u"贸易商资源")
        subsheet.write(rowindex, 16 + 6 * (len(powder) + len(block)), u"贸易商占比")
        rowindex += 1
    # 数据部分
    subsheet.write(rowindex, 0, date)
    subsheet.write(rowindex, 1, totalrow[1])
    subsheet.write(rowindex, 2, mainrow[1])
    subsheet.write(rowindex, 3, getratio(mainrow[1], totalrow[1]))
    subsheet.write(rowindex, 4, powderrow[1])
    subsheet.write(rowindex, 5, getratio(powderrow[1], mainrow[1]))
    for i in range(0, len(powder)):
        subsheet.write(rowindex, 6 + i * 6, goodsdata[i][0])
        subsheet.write(rowindex, 7 + i * 6, getratio(goodsdata[i][0], powderrow[1]))
        subsheet.write(rowindex, 8 + i * 6, goodsdata[i][1])
        subsheet.write(rowindex, 9 + i * 6, goodsdata[i][2])
        subsheet.write(rowindex, 10 + i * 6, goodsdata[i][3])
        subsheet.write(rowindex, 11 + i * 6, goodsdata[i][4])
    subsheet.write(rowindex, 6 + 6 * len(powder), blockrow[1])
    subsheet.write(rowindex, 7 + 6 * len(powder), getratio(blockrow[1], mainrow[1]))
    for i in range(len(powder), (len(powder)+len(block))):
        subsheet.write(rowindex, 8 + 6 * i, goodsdata[i][0])
        subsheet.write(rowindex, 9 + 6 * i, getratio(goodsdata[i][0], blockrow[1]))
        subsheet.write(rowindex, 10 + 6 * i, goodsdata[i][1])
        subsheet.write(rowindex, 11 + 6 * i, goodsdata[i][2])
        subsheet.write(rowindex, 12 + 6 * i, goodsdata[i][3])
        subsheet.write(rowindex, 13 + 6 * i, goodsdata[i][4])
    subsheet.write(rowindex, 8 + 6 * (len(powder) + len(block)), mainrow[2])
    subsheet.write(rowindex, 9 + 6 * (len(powder) + len(block)), mainrow[3])
    subsheet.write(rowindex, 10 + 6 * (len(powder) + len(block)), mainrow[4])
    subsheet.write(rowindex, 11 + 6 * (len(powder) + len(block)), mainrow[5])
    subsheet.write(rowindex, 12 + 6 * (len(powder) + len(block)), nonmainrow[1])
    subsheet.write(rowindex, 13 + 6 * (len(powder) + len(block)), nonmainrow[2])
    subsheet.write(rowindex, 14 + 6 * (len(powder) + len(block)), nonmainrow[3])
    subsheet.write(rowindex, 15 + 6 * (len(powder) + len(block)), nonmainrow[4])
    subsheet.write(rowindex, 16 + 6 * (len(powder) + len(block)), nonmainrow[5])

    return trackfile