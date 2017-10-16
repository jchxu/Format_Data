# coding=utf-8
import xlrd, xlwt, re
from os import path
from Read_Data import *
from xlutils.display import *

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

# 设置背景颜色
pattern = xlwt.Pattern()
pattern.pattern = xlwt.Pattern.SOLID_PATTERN
pattern.pattern_fore_colour = 5
# 设置对齐方式
alignmentCenter = xlwt.Alignment()
alignmentCenter.horz = xlwt.Alignment.HORZ_CENTER
alignmentCenter.vert = xlwt.Alignment.VERT_CENTER
alignmentLeft = xlwt.Alignment()
alignmentLeft.horz = xlwt.Alignment.HORZ_LEFT
alignmentLeft.vert = xlwt.Alignment.VERT_CENTER
alignmentRight = xlwt.Alignment()
alignmentRight.horz = xlwt.Alignment.HORZ_RIGHT
alignmentRight.vert = xlwt.Alignment.VERT_CENTER
# 设置字体样式
fontB = xlwt.Font()
fontB.bold = True


# 创建格式
#style = xlwt.XFStyle()
#style.pattern = pattern
#style.alignment = alignment


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
    print u"货主、货种的名称已替换为标准名称."
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
    if numa != 0.0 and numb != 0.0 :
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

def calculate_summary(owner, goods, amount, company, goods_class_list, goods_class_name):
    ### 计算所有货物总和 ###
    totalamount = sum_owner_goods(owner, goods, amount, u"ALL", u"ALL")
    totalcom = 0
    totaltrader = 0
    for item in set(owner.values()):
        if item in company:
            totalcom += sum_owner_goods(owner, goods, amount, item, u"ALL")
        else:
            totaltrader += sum_owner_goods(owner, goods, amount, item, u"ALL")
    totalcomratio = getratio(totalcom, totalamount)
    totaltraderratio = getratio(totaltrader, totalamount)
    #print totalamount, totalcom, totalcomratio, totaltrader, totaltraderratio
    totalrow = [u"港存总计", totalamount, totalcom, totalcomratio, totaltrader, totaltraderratio]

    ### 计算每种货物总和 ###
    goodstotal = {}
    goodscom = {}
    goodscomratio = {}
    goodstrader = {}
    goodstraderratio = {}
    goodsrow = {}
    biglist = []
    for i in range(0, len(goods_class_name)):
        for k in range(0, len(goods_class_list[i])):
            biglist.append(goods_class_list[i][k])
    for i in range(0, len(biglist)):
        out = getgoodsamount(biglist[i], owner, goods, amount, company)
        goodstotal[i], goodscom[i], goodscomratio[i], goodstrader[i], goodstraderratio[i] = out
        #print goodstotal[i], goodscom[i], goodscomratio[i], goodstrader[i], goodstraderratio[i]
        goodsrow[i] = [biglist[i], goodstotal[i], goodscom[i], goodscomratio[i], goodstrader[i], goodstraderratio[i]]

    ### 分别统计不同大类情况 ###
    class_total = {}
    class_com = {}
    class_trader = {}
    class_comratio = {}
    class_traderratio = {}
    classrow = {}
    end = len(goods_class_name) -1
    class_total[end] = totalamount
    class_com[end] = totalcom
    class_trader[end] = totaltrader
    for i in range(0, len(goods_class_name)-1):
        index = 0
        if i > 0:
            for x in range(0, i):
                index += len(goods_class_list[x])
        class_total[i] = 0
        class_com[i] = 0
        class_trader[i] = 0
        for k in range(0, len(goods_class_list[i])):
            class_total[i] += goodstotal[index+k]
            class_com[i] += goodscom[index+k]
            class_trader[i] += goodstrader[index+k]
        class_comratio[i] = getratio(class_com[i], class_total[i])
        class_traderratio[i] = getratio(class_trader[i], class_total[i])
        classrow[i] = [goods_class_name[i], class_total[i], class_com[i], class_comratio[i], class_trader[i], class_traderratio[i]]
        # 最后一个大类（非主流资源）的数据为总量-前面所有分类的量
        class_total[end] -= class_total[i]
        class_com[end] -= class_com[i]
        class_trader[end] -= class_trader[i]
    class_comratio[end] = getratio(class_com[end], class_total[end])
    class_traderratio[end] = getratio(class_trader[end], class_total[end])
    classrow[end] = [goods_class_name[end], class_total[end], class_com[end], class_comratio[end], class_trader[end], class_traderratio[end]]

    ### 计算主流资源（分类列表中的前两个） ###
    mainamount = 2
    maintotal = 0
    maincom = 0
    maintrader = 0
    for i in range(0, mainamount):
        maintotal += class_total[i]
        maincom += class_com[i]
        maintrader += class_trader[i]
    maincomratio = getratio(maincom, maintotal)
    maintraderratio = getratio(maintrader, maintotal)
    #print maintotal, maincom, maincomratio, maintrader, maintraderratio
    mainrow = [u"主流资源", maintotal, maincom, maincomratio, maintrader, maintraderratio]

    return (totalrow, mainrow, classrow, goodsrow)

def summary_style(level, num):
    if level == 'total' and num == 0:
        style = xlwt.easyxf("font: bold on, color-index blue; alignment: vert center, horz left; pattern: pattern solid, fore_colour light_yellow;")
    elif level == 'total' and num in [1, 2, 4]:
        style = xlwt.easyxf("font: bold on, color-index blue; alignment: vert center, horz right; pattern: pattern solid, fore_colour light_yellow;", num_format_str='#,##0')
    elif level == 'total' and num in [3, 5]:
        style = xlwt.easyxf("font: bold on, color-index blue; alignment: vert center, horz center; pattern: pattern solid, fore_colour light_yellow;", num_format_str='0.00%')
    elif level == 'title1' and num == 0:
        style = xlwt.easyxf("font: bold off; alignment: vert center, horz left; pattern: pattern solid, fore_colour ice_blue;")
    elif level == 'title1' and num in [1,2,4]:
        style = xlwt.easyxf("font: bold off; alignment: vert center, horz right; pattern: pattern solid, fore_colour ice_blue;", num_format_str='#,##0')
    elif level == 'title1' and num in [3,5]:
        style = xlwt.easyxf("font: bold off; alignment: vert center, horz center; pattern: pattern solid, fore_colour ice_blue;", num_format_str='0.00%')
    elif level == 'title2' and num == 0:
        style = xlwt.easyxf("font: bold on, color-index white; alignment: vert center, horz left; pattern: pattern solid, fore_colour light_blue;")
    elif level == 'title2' and num in [1,2,4]:
        style = xlwt.easyxf("font: bold on, color-index white; alignment: vert center, horz right; pattern: pattern solid, fore_colour light_blue;", num_format_str='#,##0')
    elif level == 'title2' and num in [3,5]:
        style = xlwt.easyxf("font: bold on, color-index white; alignment: vert center, horz center; pattern: pattern solid, fore_colour light_blue;", num_format_str='0.00%')
    elif level == 'goods' and num == 0:
        style = xlwt.easyxf("alignment: vert center, horz left;")
    elif level == 'goods' and num in [1,2,4]:
        style = xlwt.easyxf("alignment: vert center, horz right;", num_format_str='#,##0')
    elif level == 'goods' and num in [3,5]:
        style = xlwt.easyxf("alignment: vert center, horz center;", num_format_str='0.00%')
    return style

def write_summary(resultfile, goods_class_list, totalrow, mainrow, classrow, goodsrow):
    subsheet = resultfile.add_sheet("Summary")
    titlerow = [u"矿种", u"合计", u"钢厂", u"钢厂占比", u"贸易商", u"贸易商占比"]
    for i in range(0, len(titlerow)):
        subsheet.write(0, i, titlerow[i], xlwt.easyxf("font: bold on; alignment: vert center, horz center;"))
        subsheet.write(1, i, totalrow[i], summary_style('total', i))
        subsheet.write(2, i, mainrow[i], summary_style('title2', i))
    for k in range(0, len(classrow)):
        index = 3   # 仅比mainrow大1，表示紧挨着下一行写数据；若有空1行，则index=4
        index2 = 0
        if k > 0:
            for x in range(0, k):
                index += (len(goods_class_list[x])+2)
                index2 += len(goods_class_list[x])
        for m in range(0, len(titlerow)):
            if k != len(classrow)-1:
                subsheet.write(index, m, classrow[k][m], summary_style('title1', m))
            else:
                subsheet.write(index, m, classrow[k][m], summary_style('title2', m))
        for n in range(0, len(goods_class_list[k])):
            for y in range(0, len(titlerow)):
                subsheet.write(index+n+1, y, goodsrow[index2+n][y], summary_style('goods', y))

    print u'统计信息已写入子表"\033[1;34;0m%s\033[0m".' % subsheet.name.encode('utf-8')
    return resultfile

def write_detail(resultfile, owner, goods, amount):
    """按照序号、货主、货物、数量的顺序写入所有数据"""
    ### 设置输出格式 ###
    style_title = xlwt.easyxf("font: bold on, color-index blue; alignment: vert center, horz center; pattern: pattern solid, fore_colour light_yellow;")
    style_center = xlwt.easyxf("alignment: vert center, horz center;")
    style_name = xlwt.easyxf("alignment: vert center, horz left;")
    style_amount = xlwt.easyxf("alignment: vert center, horz right;", num_format_str='#,##0')

    subsheet = resultfile.add_sheet("Detail")
    titlerow = [u"序号", u"货主", u"货名", u"数量"]
    ### 写标题行 ###
    for i in range(0, len(titlerow)):
        subsheet.write(0, i, titlerow[i],style_title)
    ### 写每行数据 ###
    for i in range(1, len(owner.keys()) + 1):
        subsheet.write(i, 0, i, style_center)
        subsheet.write(i, 1, owner[i], style_name)
        subsheet.write(i, 2, goods[i], style_name)
        subsheet.write(i, 3, amount[i], style_amount)
    print u'详细信息已写入子表"\033[1;34;0m%s\033[0m".' % subsheet.name.encode('utf-8')
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

def set_row(index):
    dic = {}
    if index == 1:
        return 2
    elif index in [2,3,4,5]:
        return index+2

def track_style(num):
    if num in [2,4]:
        style = xlwt.easyxf("alignment: vert center, horz right;", num_format_str='#,##0')
    elif num in [3,5]:
        style = xlwt.easyxf("alignment: vert center, horz center;", num_format_str='0.00%')
    return style

def write_tracking(tracklist, stddate, olddate, trackfile, subsheet, rowindex, goods_class_name, goods_class_list, totalrow, mainrow, classrow, goodsrow):
    """追加输出历史追踪数据"""
    # 判断是否有追踪数据的品种清单。若有，则按清单追踪；如无，则追踪所有大类品种。
    mainamount = 2
    style_num = xlwt.easyxf("alignment: vert center, horz right;", num_format_str='#,##0')
    style_precent = xlwt.easyxf("alignment: vert center, horz center;", num_format_str='0.00%')
    style_title = xlwt.easyxf("font: bold on; alignment: vert center, horz left;")
    #if path.exists(tracklist.decode('utf-8')):
    #    print
    #else:
    #    print

    # 数据部分
    if rowindex == 0:
        writeindex = 5
    else:
        writeindex = rowindex
    year = get_date_time()[1]
    month = stddate[0:2]
    day = stddate[2:4]
    date = "%4s/%2s/%2s" % (year, month, day)
    #print olddate
    #for i in olddate.keys():
    #    if date == olddate[i]:
    #        writeindex = i
    subsheet.write(writeindex, 0, date)
    subsheet.write(writeindex, 1, totalrow[1], style_num)
    for i in range(1, 6):
        if i == 1:
            subsheet.write(writeindex, set_row(i), mainrow[i], style_num)
            subsheet.write(writeindex, set_row(i)+1, getratio(mainrow[1], totalrow[1]), style_precent)
        else:
            subsheet.write(writeindex, set_row(i), mainrow[i], track_style(i))
    for i in range(0, len(classrow)):
        index = 8
        index2 = 0
        if i > 0:
            for x in range(0, i):
                index += (len(goods_class_list[x])*6)+1
                index2 += len(goods_class_list[x])
        if i < mainamount:
            subsheet.write(writeindex, index+i, classrow[i][1], style_num)
            subsheet.write(writeindex, index+i+1, getratio(classrow[i][1], mainrow[1]), style_precent)
        elif i != len(classrow)-1:
            subsheet.write(writeindex, index+i, classrow[i][1], style_num)
            subsheet.write(writeindex, index+i+1, getratio(classrow[i][1], totalrow[1]), style_precent)
        else:
            subsheet.write(writeindex, index+i, classrow[i][1], style_num)
            subsheet.write(writeindex, index+i+1, getratio(classrow[i][1], totalrow[1]), style_precent)
        for k in range(0, len(goods_class_list[i])):
            subsheet.write(writeindex, index+i+2+(6*k), goodsrow[index2+k][1], style_num)
            subsheet.write(writeindex, index+i+2+(6*k)+1, getratio(goodsrow[index2+k][1], classrow[i][1]), style_precent)
            subsheet.write(writeindex, index+i+2+(6*k)+2, goodsrow[index2+k][2], style_num)
            subsheet.write(writeindex, index+i+2+(6*k)+3, goodsrow[index2+k][3], style_precent)
            subsheet.write(writeindex, index+i+2+(6*k)+4, goodsrow[index2+k][4], style_num)
            subsheet.write(writeindex, index+i+2+(6*k)+5, goodsrow[index2+k][5], style_precent)

        # 标题部分
    titleindex = 4
    subsheet.write_merge(0, titleindex, 0, 0, u"日期", style_title)
    subsheet.write_merge(0, titleindex, 1, 1, u"总库存", style_title)
    subsheet.write_merge(1, titleindex, 2, 2, mainrow[0], style_title)
    subsheet.write(titleindex, 3, u"占比", style_title)
    subsheet.write_merge(3, titleindex, 4, 4, u"钢厂", style_title)
    subsheet.write(titleindex, 5, u"钢厂占比", style_title)
    subsheet.write_merge(3, titleindex, 6, 6, u"贸易商", style_title)
    subsheet.write(titleindex, 7, u"贸易商占比", style_title)
    for i in range(0, len(classrow)):
        index = 8  # 前8列标题已指定
        index2 = 0
        if i > 0:
            for x in range(0, i):
                index += (len(goods_class_list[x])*6)+1
                index2 += len(goods_class_list[x])
        if i != len(classrow)-1:
            subsheet.write_merge(2, titleindex, index+i, index+i, classrow[i][0], style_title)
            subsheet.write(titleindex, index+i+1, u"占比", style_title)
        else:
            subsheet.write_merge(1, titleindex, index+i, index+i, classrow[i][0], style_title)
            subsheet.write(titleindex, index+i+1, u"占比", style_title)
        for k in range(0, len(goods_class_list[i])):
            subsheet.write_merge(3, titleindex, index+i+2+(6*k), index+i+2+(6*k), goodsrow[index2+k][0], style_title)
            subsheet.write(titleindex, index+i+2+(6*k)+1, goodsrow[index2+k][0]+u"占比", style_title)
            subsheet.write(titleindex, index+i+2+(6*k)+2, u"钢厂", style_title)
            subsheet.write(titleindex, index+i+2+(6*k)+3, u"钢厂占比", style_title)
            subsheet.write(titleindex, index+i+2+(6*k)+4, u"贸易商", style_title)
            subsheet.write(titleindex, index+i+2+(6*k)+5, u"贸易商占比", style_title)

    return (trackfile, writeindex)

def calculate_sum_summary(dates, owner0, goods0, amount0, company, goods_class_list, goods_class_name):
    totalrow = {}
    mainrow = {}
    classrow = {}
    goodsrow = {}
    ### 根据不同日期分组保存数据 ###
    for index in range(0, len(dates.keys())):
        owner = {}
        goods = {}
        amount = {}
        lowindex = 0
        dateitem = dates.keys()[index]
        if index > 0:
            for i in range(0, index):
                lowindex += dates[dates.keys()[i]]
        #print lowindex, lowindex+dates[dateitem]
        for i in range(lowindex, lowindex+dates[dateitem]):
            owner[i-lowindex] = owner0[i]
            goods[i-lowindex] = goods0[i]
            amount[i-lowindex] = amount0[i]
        ### 计算所有货物总和 ###
        totalamount = sum_owner_goods(owner, goods, amount, u"ALL", u"ALL")
        totalcom = 0
        totaltrader = 0
        for item in set(owner.values()):
            if item in company:
                totalcom += sum_owner_goods(owner, goods, amount, item, u"ALL")
            else:
                totaltrader += sum_owner_goods(owner, goods, amount, item, u"ALL")
        totalcomratio = getratio(totalcom, totalamount)
        totaltraderratio = getratio(totaltrader, totalamount)
        #print totalamount, totalcom, totalcomratio, totaltrader, totaltraderratio
        totalrow[dateitem] = [u"港存总计", totalamount, totalcom, totalcomratio, totaltrader, totaltraderratio]
        ### 计算每种货物总和 ###
        goodstotal = {}
        goodscom = {}
        goodscomratio = {}
        goodstrader = {}
        goodstraderratio = {}
        goodsrow0 = {}
        biglist = []
        for i in range(0, len(goods_class_name)):
            for k in range(0, len(goods_class_list[i])):
                biglist.append(goods_class_list[i][k])
        for i in range(0, len(biglist)):
            out = getgoodsamount(biglist[i], owner, goods, amount, company)
            goodstotal[i], goodscom[i], goodscomratio[i], goodstrader[i], goodstraderratio[i] = out
            #print goodstotal[i], goodscom[i], goodscomratio[i], goodstrader[i], goodstraderratio[i]
            goodsrow0[i] = [biglist[i], goodstotal[i], goodscom[i], goodscomratio[i], goodstrader[i], goodstraderratio[i]]
        goodsrow[dateitem] = goodsrow0
        ### 分别统计不同大类情况 ###
        class_total = {}
        class_com = {}
        class_trader = {}
        class_comratio = {}
        class_traderratio = {}
        classrow0 = {}
        end = len(goods_class_name) -1
        class_total[end] = totalamount
        class_com[end] = totalcom
        class_trader[end] = totaltrader
        for i in range(0, len(goods_class_name)-1):
            index = 0
            if i > 0:
                for x in range(0, i):
                    index += len(goods_class_list[x])
            class_total[i] = 0
            class_com[i] = 0
            class_trader[i] = 0
            for k in range(0, len(goods_class_list[i])):
                class_total[i] += goodstotal[index+k]
                class_com[i] += goodscom[index+k]
                class_trader[i] += goodstrader[index+k]
            class_comratio[i] = getratio(class_com[i], class_total[i])
            class_traderratio[i] = getratio(class_trader[i], class_total[i])
            classrow0[i] = [goods_class_name[i], class_total[i], class_com[i], class_comratio[i], class_trader[i], class_traderratio[i]]
            # 最后一个大类（非主流资源）的数据为总量-前面所有分类的量
            class_total[end] -= class_total[i]
            class_com[end] -= class_com[i]
            class_trader[end] -= class_trader[i]
        class_comratio[end] = getratio(class_com[end], class_total[end])
        class_traderratio[end] = getratio(class_trader[end], class_total[end])
        classrow0[end] = [goods_class_name[end], class_total[end], class_com[end], class_comratio[end], class_trader[end], class_traderratio[end]]
        classrow[dateitem] = classrow0
        ### 计算主流资源（分类列表中的前两个） ###
        mainamount = 2
        maintotal = 0
        maincom = 0
        maintrader = 0
        for i in range(0, mainamount):
            maintotal += class_total[i]
            maincom += class_com[i]
            maintrader += class_trader[i]
        maincomratio = getratio(maincom, maintotal)
        maintraderratio = getratio(maintrader, maintotal)
        #print maintotal, maincom, maincomratio, maintrader, maintraderratio
        mainrow[dateitem] = [u"主流资源", maintotal, maincom, maincomratio, maintrader, maintraderratio]
    return (totalrow, mainrow, classrow, goodsrow)

def write_sum_summary(resultfile, dateitem, goods_class_list, totalrow, mainrow, classrow, goodsrow):
    subsheet = resultfile.add_sheet("Summary")
    titlerow = [u"矿种", u"合计", u"钢厂", u"钢厂占比", u"贸易商", u"贸易商占比"]
    for i in range(0, len(titlerow)):
        subsheet.write(0, i, titlerow[i], xlwt.easyxf("font: bold on; alignment: vert center, horz center;"))
        subsheet.write(1, i, totalrow[i], summary_style('total', i))
        subsheet.write(2, i, mainrow[i], summary_style('title2', i))
    for k in range(0, len(classrow)):
        index = 3   # 仅比mainrow大1，表示紧挨着下一行写数据；若有空1行，则index=4
        index2 = 0
        if k > 0:
            for x in range(0, k):
                index += (len(goods_class_list[x])+2)
                index2 += len(goods_class_list[x])
        for m in range(0, len(titlerow)):
            if k != len(classrow)-1:
                subsheet.write(index, m, classrow[k][m], summary_style('title1', m))
            else:
                subsheet.write(index, m, classrow[k][m], summary_style('title2', m))
        for n in range(0, len(goods_class_list[k])):
            for y in range(0, len(titlerow)):
                subsheet.write(index+n+1, y, goodsrow[index2+n][y], summary_style('goods', y))
    print u'\033[1;34;0m%s\033[0m各港口汇总统计数据已写入Summary子表.' % dateitem
    return resultfile

def write_sum_detail(resultfile, item, dates, port, owner, goods, amount):
    style_title = xlwt.easyxf("font: bold on, color-index blue; alignment: vert center, horz center; pattern: pattern solid, fore_colour light_yellow;")
    style_center = xlwt.easyxf("alignment: vert center, horz center;")
    style_name = xlwt.easyxf("alignment: vert center, horz left;")
    style_amount = xlwt.easyxf("alignment: vert center, horz right;", num_format_str='#,##0')

    subsheet = resultfile.add_sheet("Detail")
    titlerow = [u"序号", u"港口", u"货主", u"货名", u"数量"]
    ### 写标题行 ###
    for i in range(0, len(titlerow)):
        subsheet.write(0, i, titlerow[i], style_title)
    ### 写每行数据 ###
    count = 0
    dateindex = dates.keys().index(item)
    if dateindex > 0:
        for i in range(0, dateindex):
            count += dates[dates.keys()[i]]
    for i in range(0, dates[item]):
        subsheet.write(i + 1, 0, i + 1, style_center)
        subsheet.write(i + 1, 1, port[count], style_center)
        subsheet.write(i + 1, 2, owner[count], style_name)
        subsheet.write(i + 1, 3, goods[count], style_name)
        subsheet.write(i + 1, 4, amount[count], style_amount)
        count += 1
    print u'\033[1;34;0m%s\033[0m各港口汇总共计\033[1;34;0m%s\033[0m条数据已写入Detail子表.' % (item, dates[item])
    return resultfile

def write_sum_tracking(dateitem, trackfile, subsheet, rowindex, goods_class_name, goods_class_list, totalrow, mainrow, classrow, goodsrow):
    """追加输出历史追踪数据"""
    mainamount = 2
    style_num = xlwt.easyxf("alignment: vert center, horz right;", num_format_str='#,##0')
    style_precent = xlwt.easyxf("alignment: vert center, horz center;", num_format_str='0.00%')
    style_title = xlwt.easyxf("font: bold on; alignment: vert center, horz left;")
    # 数据部分
    if rowindex == 0:
        writeindex = 5
    else:
        writeindex = rowindex
    date = get_date_time()[1]+dateitem
    subsheet.write(writeindex, 0, date)
    subsheet.write(writeindex, 1, totalrow[1], style_num)
    for i in range(1, 6):
        if i == 1:
            subsheet.write(writeindex, set_row(i), mainrow[i], style_num)
            subsheet.write(writeindex, set_row(i)+1, getratio(mainrow[1], totalrow[1]), style_precent)
        else:
            subsheet.write(writeindex, set_row(i), mainrow[i], track_style(i))
    for i in range(0, len(classrow)):
        index = 8
        index2 = 0
        if i > 0:
            for x in range(0, i):
                index += (len(goods_class_list[x])*6)+1
                index2 += len(goods_class_list[x])
        if i < mainamount:
            subsheet.write(writeindex, index+i, classrow[i][1], style_num)
            subsheet.write(writeindex, index+i+1, getratio(classrow[i][1], mainrow[1]), style_precent)
        elif i != len(classrow)-1:
            subsheet.write(writeindex, index+i, classrow[i][1], style_num)
            subsheet.write(writeindex, index+i+1, getratio(classrow[i][1], totalrow[1]), style_precent)
        else:
            subsheet.write(writeindex, index+i, classrow[i][1], style_num)
            subsheet.write(writeindex, index+i+1, getratio(classrow[i][1], totalrow[1]), style_precent)
        for k in range(0, len(goods_class_list[i])):
            subsheet.write(writeindex, index+i+2+(6*k), goodsrow[index2+k][1], style_num)
            subsheet.write(writeindex, index+i+2+(6*k)+1, getratio(goodsrow[index2+k][1], classrow[i][1]), style_precent)
            subsheet.write(writeindex, index+i+2+(6*k)+2, goodsrow[index2+k][2], style_num)
            subsheet.write(writeindex, index+i+2+(6*k)+3, goodsrow[index2+k][3], style_precent)
            subsheet.write(writeindex, index+i+2+(6*k)+4, goodsrow[index2+k][4], style_num)
            subsheet.write(writeindex, index+i+2+(6*k)+5, goodsrow[index2+k][5], style_precent)

        # 标题部分
    titleindex = 4
    subsheet.write_merge(0, titleindex, 0, 0, u"日期", style_title)
    subsheet.write_merge(0, titleindex, 1, 1, u"总库存", style_title)
    subsheet.write_merge(1, titleindex, 2, 2, mainrow[0], style_title)
    subsheet.write(titleindex, 3, u"占比", style_title)
    subsheet.write_merge(3, titleindex, 4, 4, u"钢厂", style_title)
    subsheet.write(titleindex, 5, u"钢厂占比", style_title)
    subsheet.write_merge(3, titleindex, 6, 6, u"贸易商", style_title)
    subsheet.write(titleindex, 7, u"贸易商占比", style_title)
    for i in range(0, len(classrow)):
        index = 8  # 前8列标题已指定
        index2 = 0
        if i > 0:
            for x in range(0, i):
                index += (len(goods_class_list[x])*6)+1
                index2 += len(goods_class_list[x])
        if i != len(classrow)-1:
            subsheet.write_merge(2, titleindex, index+i, index+i, classrow[i][0], style_title)
            subsheet.write(titleindex, index+i+1, u"占比", style_title)
        else:
            subsheet.write_merge(1, titleindex, index+i, index+i, classrow[i][0], style_title)
            subsheet.write(titleindex, index+i+1, u"占比", style_title)
        for k in range(0, len(goods_class_list[i])):
            subsheet.write_merge(3, titleindex, index+i+2+(6*k), index+i+2+(6*k), goodsrow[index2+k][0], style_title)
            subsheet.write(titleindex, index+i+2+(6*k)+1, goodsrow[index2+k][0]+u"占比", style_title)
            subsheet.write(titleindex, index+i+2+(6*k)+2, u"钢厂", style_title)
            subsheet.write(titleindex, index+i+2+(6*k)+3, u"钢厂占比", style_title)
            subsheet.write(titleindex, index+i+2+(6*k)+4, u"贸易商", style_title)
            subsheet.write(titleindex, index+i+2+(6*k)+5, u"贸易商占比", style_title)
    print u'\033[1;34;0m%s\033[0m各港口汇总历史追踪数据已写入第\033[1;34;0m%d\033[0m行.' % (dateitem, writeindex + 1)
    return trackfile

def sum_by_traderandgoods(company, trader, onlyowner, onlygoods, owner, goods, amount):
    traderorder = {}  #贸易商或品种名称为key，加和后的数值为value
    goodsorder = {}
    for item in onlyowner:
        if item in trader:
            traderorder[item] = 0
    for item in onlygoods:
        goodsorder[item] = 0
    for i in range(0, len(amount)):
        if owner[i] in trader:
            traderorder[owner[i]] += amount[i]
        elif owner[i] not in company:
            print u'请检查货主名，未找到"\033[1;31;0m%s\033[0m".' % owner[i]
    for i in range(0, len(amount)):
        goodsorder[goods[i]] += amount[i]
    traderorder = sorted(traderorder.iteritems(), key=lambda d: d[1], reverse=True)
    goodsorder = sorted(goodsorder.iteritems(), key=lambda d: d[1], reverse=True)
    return traderorder, goodsorder

def write_by_traderandgoods(ownershipfile, item, dates, port, owner, goods, amount):
    style_title = xlwt.easyxf("font: bold on, color-index blue; alignment: vert center, horz center; pattern: pattern solid, fore_colour light_yellow;")
    style_center = xlwt.easyxf("alignment: vert center, horz center;")
    style_name = xlwt.easyxf("alignment: vert center, horz left;")
    style_amount = xlwt.easyxf("alignment: vert center, horz right;", num_format_str='#,##0')

    subsheet = ownershipfile.add_sheet("Detail")
    titlerow = [u"排序", u"贸易商", u"数量", u"   ", u"排序", u"品种", u"数量"]
    ### 写标题行 ###
    for i in range(0, len(titlerow)):
        if i != 3:
            subsheet.write(1, i, titlerow[i], style_title)

