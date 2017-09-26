# -*- coding:utf8 -*-
import pymssql
import xlsxwriter as ws
import datetime
import pandas
import numpy
import math

# 质量选股模型
server = '192.168.8.200'
user = 'sa'
password = 'wind123@pa'
database = 'PortfolioData'
conn = pymssql.connect(server, user, password, database, charset='utf8')
# HS300的成分股及行业、市值、权重sql
# sqlHS300Industry = '''with t0 as (select top 1 FDate from IndexConstituent where FDate < %s order by FDate desc)
# select a.FDate, b.SecCode, b.IndustryName1, b.IndustryCode1, c.FreeFloatShare * d.AdjClose mv, a.[Weight] / 100 as [Weight] from t0, IndexConstituent a,StockIndustry b, StockShareInfo c, SecPrice d where substring(a.SecCode, 1, 6) = b.SecCode and a.FDate >= b.StartDate and a.FDate <= ISNULL(b.EndDate, '2050-12-31') and b.IndustryType = 1 and a.FDate = d.FDate and a.SecCode = d.SecCode and a.FDate >= c.StartDate and a.FDate <= isnull(c.EndDate, '2050-12-31') and b.SecCode = c.SecCode and a.FDate = t0.FDate and a.IndexCode = '000300.SH'
#  '''
sqlHS300Industry = '''with t0 as (select top 1 FDate from IndexConstituent where FDate < %s order by FDate desc)
select a.FDate, b.SecCode, e.mshiname, e.mshi, c.FreeFloatShare * d.AdjClose mv, a.[Weight] / 100 as [Weight] from t0, IndexConstituent a,StockIndustry b, StockShareInfo c, SecPrice d, ci3mshinew e where substring(a.SecCode, 1, 6) = b.SecCode and a.FDate >= b.StartDate and a.FDate <= ISNULL(b.EndDate, '2050-12-31') and b.IndustryType = 3 and a.FDate = d.FDate and a.SecCode = d.SecCode and a.FDate >= c.StartDate and a.FDate <= isnull(c.EndDate, '2050-12-31') and b.SecCode = c.SecCode and a.FDate = t0.FDate and a.IndexCode = '000300.SH' and e.IndustryName3 = b.IndustryName3 '''
sqlPctChg = '''select a.FDate, b.SecCode2, a.PctChg from StockMovingAvg a, SecInfo b where a.SecCode = b.SecCode and b.SecType = 'A' and a.FDate > (select top 1 t.[Day] from TradeDay t where t.[Day] > '%s' order by 1) and a.FDate <= (select top 1 t.[Day] from TradeDay t where t.[Day] > '%s' order by 1) and b.SecCode2 in(%s) order by 1,2'''
sqlHS300PctChg = '''select a.FDate, a.PctChg from StockMovingAvg a where a.SecCode = '000300.SH' and a.FDate > (select top 1 t.[Day] from TradeDay t where t.[Day] > %s order by 1) and a.FDate <= (select top 1 t.[Day] from TradeDay t where t.[Day] > %s order by 1) '''
cursor = conn.cursor()
# 执行选股收益率统计sql
today = datetime.datetime.now().strftime('%Y%m%d')
daylist = ('-04-30', '-08-30')
sqldata1 = []
dayReturn = []
codeCount = []
startyear = 2006
k = 0
while k <= 10:
    year = startyear + k
    for day in daylist:
        startday = str(year) + day
        if day == '-04-30':
            endDay = str(year) + '-08-30'
        else:
            endDay = str(year + 1) + '-04-30'
        cursor.execute('exec [GetNewChooseCode] @fDate = %s', startday)
        data1 = cursor.fetchall()
        cursor.execute(sqlHS300Industry, startday)
        data2 = cursor.fetchall()
        dataDict1 = {'IndustryCode': map(lambda x: x[4], data1), 'MV': map(lambda x: x[5], data1)}
        dataDict2 = {'IndustryCode': map(lambda x: x[3], data2), 'MV': map(lambda x: x[4], data2), 'Weight': map(lambda x: x[5], data2)}
        dataFrameDict1 = pandas.DataFrame(dataDict1)
        dataFrameDict2 = pandas.DataFrame(dataDict2)
        dict1 = dataFrameDict1.groupby(dataFrameDict1['IndustryCode']).sum()
        dict2 = dataFrameDict2.groupby(dataFrameDict2['IndustryCode']).sum()
        count1 = dataFrameDict1.groupby(dataFrameDict1['IndustryCode']).size()
        count2 = dataFrameDict2.groupby(dataFrameDict2['IndustryCode']).size()
        # 比较HS300和质量篮子的行业分布{行业代码:[股票家数,行业总市值,行业权重]}
        aList = {a: [count1[a], dict1['MV'][a]] for a in list(dict1['MV'].keys()) if a not in list(dict2['MV'].keys())}  # a.存在于质量篮子但不存在于HS300篮子,保留质量篮子行业,权重未计算
        # bList = {b: [count1[b], dict1['MV'][b], dict2['Weight'][b]] for b in list(dict1['MV'].keys()) if b in list(dict2['MV'].keys())}     # b.质量和HS300交集部分,将HS300行业权重分配给质量
        bList = {b: [count1[b], dict2['MV'][b], dict2['Weight'][b]] for b in list(dict1['MV'].keys()) if b in list(dict2['MV'].keys())}     # b.质量和HS300交集部分,将HS300行业权重分配给质量
        cList = {c: [count2[c], dict2['MV'][c], dict2['Weight'][c]] for c in list(dict2['MV'].keys()) if c not in list(dict1['MV'].keys())}   # c.存在HS300但不存在质量篮子,将HS300行业及权重添加到质量
        # 根据所有行业市值计算a篮子的权重
        aSum = sum([j[1] for j in aList.values()])
        bSum = sum([j[1] for j in bList.values()])
        cSum = sum([j[1] for j in cList.values()])
        for j in aList.values():
            j.append(j[1] / (aSum + bSum + cSum))
        for j in bList.values():
            j[2] *= (bSum + cSum) / (aSum + bSum + cSum)
        for j in cList.values():
            j[2] = (bSum + cSum) / (aSum + bSum + cSum)
        bList.update(aList)
        # 整合筛选出来的代码及对应权重
        codeDict = {}
        for i in range(0, len(data1)):
            if data1[i][4] in bList.keys():
                codeDict[data1[i][2]] = bList[data1[i][4]][2] / bList[data1[i][4]][0]
        for m in range(0, len(data2)):
            if data2[m][3] in cList.keys():
                codeDict[data2[m][1]] = cList[data2[m][3]][2] * data2[m][5]
        # 计算收益率
        codeList = codeDict.keys()
        codeString = ('%s,' * len(codeList))[:-1]
        sqlPctChg1 = sqlPctChg %(startday, endDay, codeString)
        cursor.execute(sqlPctChg1, tuple(codeList))
        data3 = cursor.fetchall()
        newData3 = []
        for n in range(0, len(data3)):
            newData3.append([data3[n][0], data3[n][1], data3[n][2] * codeDict[data3[n][1]]])
        codePctChg = {'FDate': map(lambda x: x[0], newData3), 'PctChg': map(lambda x: x[2], newData3)}
        codeData = pandas.DataFrame(codePctChg, dtype=float)
        codeReturn = codeData['PctChg'].groupby(codeData['FDate']).sum() - 0.003 * 2 / 240
        # 查询HS300的每日收益率
        cursor.execute(sqlHS300PctChg, (startday, endDay))
        data4 = cursor.fetchall()
        indexReturn = dict(data4)
        sqldata1.extend(map(lambda x: [startday, x], codeList))
        dayReturn.extend(map(lambda x: [x, codeReturn[x], indexReturn[x], float(codeReturn[x]) - float(indexReturn[x])], codeReturn.keys()))
        codeCount.append([startday, int(sum([l[0] for l in aList.values()])), int(sum([l[0] for l in bList.values()])), int(sum([l[0] for l in cList.values()]))])
    k += 1
dayReturn.sort(key=lambda x: x[0])
# 计算单位净值，收益率，风险率，收益风险比，年化收益率
dayReturn[0].append(1 * (1 + dayReturn[0][3]))
for p in range(1, len(dayReturn)):
    dayReturn[p].append(dayReturn[p - 1][4] * (1 + dayReturn[p][3]))
groupReturn = sum([q[3] for q in dayReturn]) / len(dayReturn)
riskProportion = numpy.std([q[3] for q in dayReturn], ddof=1)
riskReturn = groupReturn / riskProportion * math.sqrt(240)
yearReturn = groupReturn * 240
allList = [groupReturn, riskProportion, riskReturn, yearReturn]
# 将sql执行结果插入excle
workbook = ws.Workbook('C:\Users\Administrator\Desktop\Quality_HS300_%s.xlsx' % today)
sheet1 = workbook.add_worksheet(u'调仓清单')
sheet2 = workbook.add_worksheet(u'收益计算')
field1 = [u'调仓日', u'证券代码']
field2 = [u'交易日', u'模拟组合', u'沪深300', u'超额收益', u'单位净值']
field3 = ['FDate', 'Quality', 'QH', 'HS300']
field4 = [u'收益率', u'风险波动率', u'收益风险比', u'年化收益率']
for i in range(0, len(field1)):
    sheet1.write(0, i, field1[i])
for row in range(1, len(sqldata1) + 1):
    for col in range(0, len(field1)):
        sheet1.write(row, col, sqldata1[row - 1][col])
for i in range(0, len(field2)):
    sheet2.write(0, i, field2[i])
for row in range(1, len(dayReturn) + 1):
    for col in range(0, len(field2)):
        sheet2.write(row, col, dayReturn[row - 1][col])
for i in range(0, len(field3)):
    sheet2.write(0, i + 7, field3[i])
for row in range(1, len(codeCount) + 1):
    for col in range(0, len(field3)):
        sheet2.write(row, col + 7, codeCount[row - 1][col])
for i in range(0, len(field4)):
    sheet2.write(i, 12, field4[i])
    sheet2.write(i, 13, allList[i])
sheet1.set_column('A:A', 11)
sheet2.set_column('A:A', 11)
sheet2.set_column('H:H', 11)
sheet2.set_column('M:M', 11)
chart = workbook.add_chart({'type': 'line'})
chart.add_series({
    'name': [u'收益计算', 0, 4],
    'categories': [u'收益计算', 1, 0, len(dayReturn), 0],
    'values': [u'收益计算', 1, 4, len(dayReturn), 4],
    'line': {'colorIndex': '23'},
})
chart.set_title({'name': u'单位净值'})
# chart.set_x_axis({'name': u'日期'})
# chart.set_y_axis({'name': u'单位净值'})
chart.set_size({'width': 500, 'height': 350})
sheet2.insert_chart('M7', chart)
workbook.close()
conn.close()
print 'Done!!!'
