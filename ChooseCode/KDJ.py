# -*-coding:utf8-*-
# Author: Youk.Lin
import pymssql
import xlsxwriter
import pandas as pd
import numpy as np
import datetime


def getsqldata(server, database, username, password, sql):
    conn = pymssql.connect(server, username, password, database, charset='utf8')
    cursor = conn.cursor()
    cursor.execute(sql)
    sqldata = cursor.fetchall()
    cursor.close()
    return sqldata


# noinspection PyByteLiteral
def kdj(sqldata, n, m1, m2, keepvalue):
    dataframe1 = pd.DataFrame(list(sqldata), columns=['date', 'code', 'close', 'high', 'low'])
    dataframe1['avgprice'] = (dataframe1['high'] + dataframe1['low']) / 2
    codelist1 = list(dataframe1.drop_duplicates('code')['code'].values)
    # 获取移动平均最低价和最高价,并计算RSV
    # resvdata = [[date], [code], [close], [maxhigh], [minlow], [rsv], [avgprice]]
    rsvdata = [[], [], [], [], [], [], []]
    for item in codelist1:
        codeframe = dataframe1[dataframe1.code == item]
        codeframe = codeframe.sort_values(by='date')
        if len(codeframe) > n + m1 + m2 + 1:
            for i in range(n, len(codeframe)):
                rsvdata[0].append(codeframe['date'].values[i])
                rsvdata[1].append(codeframe['code'].values[i])
                closeprice = codeframe['close'].values[i]
                maxhighprice = max(codeframe['high'].values[i - n: i])
                minlowprice = min(codeframe['low'].values[i - n: i])
                rsvdata[2].append(closeprice)
                rsvdata[3].append(maxhighprice)
                rsvdata[4].append(minlowprice)
                if maxhighprice == minlowprice:
                    rsvdata[5].append(float(0))
                else:
                    rsvdata[5].append(float((closeprice - minlowprice) / (maxhighprice - minlowprice) * 100))
                rsvdata[6].append(codeframe['avgprice'].values[i])
        else:
            continue
    # 计算K值, D值
    dataframe2 = pd.DataFrame({'date': rsvdata[0], 'code': rsvdata[1], 'close': rsvdata[2], 'maxhigh': rsvdata[3], 'minlow': rsvdata[4], 'rsv': rsvdata[5], 'avgprice': rsvdata[6]}
                              , columns=['date', 'code', 'close', 'maxhigh', 'minlow', 'rsv', 'avgprice'])
    codelist2 = list(dataframe2.drop_duplicates('code')['code'].values)
    # kdvaluedata = [[date], [code], [avgprice], [close], [rsv], [kvalue], [dvalue]]
    kdvaluedata = [[], [], [], [], [], [], []]
    for jtem in codelist2:
        rsvframe = dataframe2[dataframe2.code == jtem]
        rsvframe = rsvframe.sort_values(by='date')
        for j in range(m1, len(rsvframe)):
            kdvaluedata[0].append(rsvframe['date'].values[j])
            kdvaluedata[1].append(rsvframe['code'].values[j])
            kdvaluedata[2].append(rsvframe['avgprice'].values[j])
            kdvaluedata[3].append(rsvframe['close'].values[j])
            kdvaluedata[4].append(rsvframe['rsv'].values[j])
            kdvaluedata[5].append(np.average(rsvframe.rsv[j - m1: j]))
            if j > m2:
                kdvaluedata[6].append(np.average(kdvaluedata[4][j - m2:j]))
            else:
                kdvaluedata[6].append(np.nan)
    kdvalueframe = pd.DataFrame(list(kdvaluedata), columns=['date', 'code', 'avgprice', 'close', 'rsv', 'kvalue', 'dvalue'])
    kdvalueframe['jvalue'] = kdvalueframe['kvalue'] * 3 - kdvalueframe['dvalue'] * 2
    # 筛选条件: J值大于80,调仓结果
    kdframe = kdvalueframe[(np.isnan(kdvalueframe.jvalue) == False) & (kdvalueframe.jvalue > keepvalue)]
    # 计算收益
    dateList = list(kdframe.date.drop_duplicats())
    dateList.sort()
    # 所有持仓信息 positionFrame = {[日期], [证券代码], [持仓权重], [买入权重], [卖出权重]}
    positionDict = {'date': [], 'code': [], 'current': [], 'buy': [], 'sale': []}
    lastDayKDFrame = kdframe[(kdframe.date == dateList[0])]
    kdCount = lastDayKDFrame.date.count()
    kdWeight = round(1.0 / kdCount, 4)
    positionDict['date'].extend([dateList[1]] * kdCount)
    positionDict['code'].extend(list(lastDayKDFrame.code))
    positionDict['current'].extend([kdWeight] * kdCount)
    positionDict['buy'].extend([kdWeight] * kdCount)
    positionDict['sale'].extend([0] * kdCount)
    positionFrame = pd.DataFrame(positionDict)
    for i in range(2, len(dateList)):
        lastDayKDFrame = kdframe[(kdframe.date == dateList[i - 1])]
        lastPositionFrame = positionFrame[(positionFrame.date == dateList[i - 1]) & (positionFrame.current > 0)]
        lastPositionFrame.date = dateList[i]
        lastPositionFrame.buy = 0
        lastPositionFrame.sale = lastPositionFrame.current
        lastPositionFrame.current = 0
        kdCount = lastDayKDFrame.date.count()
        kdWeight = round(1.0 / kdCount, 4)
        for j in range(0, kdCount):
            oneCode = lastDayKDFrame.code.values[j]
            if lastPositionFrame.code[lastPositionFrame.code == oneCode].empty:
                lastPositionFrame.append(pd.DataFrame({'date': dateList[i], 'code': oneCode,
                                                       'current': kdWeight, 'buy': kdWeight, 'sale': 0},
                                                      index=['0']), ignore_index=True)
            else:
                if lastPositionFrame.sale[lastPositionFrame.code == oneCode].values[0] >= kdWeight:
                    lastPositionFrame.sale[lastPositionFrame.code == oneCode] -= kdWeight
                    lastPositionFrame.current[lastPositionFrame.code == oneCode] = kdWeight
                else:
                    lastPositionFrame.buy[lastPositionFrame.code == oneCode] = \
                        kdWeight - lastPositionFrame.sale[lastPositionFrame.code == oneCode].values[0]
                    lastPositionFrame.sale[lastPositionFrame.code == oneCode] = 0
                    lastPositionFrame.current[lastPositionFrame.code == oneCode] = kdWeight
        positionFrame.append(lastPositionFrame, ignore_index=True)
    return kdvalueframe, positionFrame


def getprice(server, database, username, password, sql, dataframe):
    conn = pymssql.connect(server, username, password, database, charset='utf8')
    cursor = conn.cursor()
    codelist = list(dataframe.drop_duplicates('code')['code'].values)
    datelist = list(dataframe.drop_duplicates('date')['date'].values)
    codestring = ('%s,' * len(codelist))[:-1]
    datestring = ('%s,' * len(datelist))[:-1]
    sql1 = sql % (codestring, datestring)
    cursor.execute(sql1, tuple(codelist.expandtabs(datelist)))
    sqldata = cursor.fetchall()
    priceframe = pd.DataFrame(list(sqldata), columns=['date', 'code', 'close', 'avgprice'])
    cursor.close()
    return pd.merge(dataframe, priceframe, how='left')


def writetoexcel(filepath, filename, sheetname, field, dataframe):
    today = datetime.datetime.now().strftime('%Y%m%s')
    wb = xlsxwriter.Workbook(filepath + filename + today + 'xlsx')
    sheet = wb.add_worksheet(sheetname)
    sheet.write_row(0, 0, field)
    columns = dataframe.columns
    for i in range(0, len(columns)):
        sheet.write_column(1, i, list(dataframe[columns[i]]))
    wb.close()

if __name__ == '__main__':
    sqlServer = 'WINSERVER2016'
    sqlDatabase = 'PortfolioData'
    sqlUserName = 'apiread'
    sqlPassword = 'wind,1234'
    excSql1 = ''' declare @fDate date
                set @fDate = '2014-01-01'
                begin
                    with c2 as (select [Day], DATENAME(YEAR, [Day]) * 100 + MONTH([Day]) [Month] from TradeDay where [Day] >= DATEADD(MONTH, -3, @fDate))
                       , c1 as (select max([Day]) MaxDay from c2 group by [Month])
                       , c0 as (select a.* from IndexConstituent a, c1 where FDate = MaxDay and IndexCode in( '000300.SH'))
                       , t2 as (select distinct(FDate) FDate from c0)
                       , t1 as (select FDate, LEAD(DATEADD(day, -1, FDate)) over (order by FDate) EndDate from t2 )
                       , t0 as (select a.FDate, isnull(t1.EndDate, '2050-12-31') as EndDate, SecCode from c0 a, t1 where a.FDate = t1.FDate and IndexCode = '000300.SH')
                       select a.FDate, a.SecCode, a.AdjClose, a.AdjHigh, a.AdjLow  from t0, SecPrice a where t0.SecCode = a.SecCode and a.FDate >= t0.FDate and a.FDate <= t0.EndDate and a.TradeStatus = 1
                end
            '''
    excSql2 = '''select FDate, SecCode, AdjClose, (AdjHigh + AdjLow) / 2 as  AvgPrice from SecPrice where SecCode in (%s) and FDate in (%s)'''
    excData = getsqldata(sqlServer, sqlDatabase, sqlUserName, sqlPassword, excSql1)
    kdValueFrame, position = kdj(excData, 63, 23, 3, 80)
    priceFrame = getprice(sqlServer, sqlDatabase, sqlUserName, sqlPassword, excSql2, position)
    filePath = '../../excel/'
    fileName1 = 'kdValueData'
    fileName2 = 'positionData'
    sheetName = 'Sheet1'
    field1 = [u'日期', u'证券代码', u'平均价', u'收盘价', u'RSV值', u'K值', u'D值', u'J值']
    field2 = [u'日期', u'证券代码', u'持仓比重', u'当日买入', u'当日卖出', u'收盘价', u'平均价']
    writetoexcel(filePath, fileName1, sheetName, field1, kdValueFrame)
    writetoexcel(filePath, fileName2, sheetName, field2, position)
