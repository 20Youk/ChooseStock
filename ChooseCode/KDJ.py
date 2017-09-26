# -*-coding:utf8-*-
# Author: Youk.Lin
import pymssql
import xlsxwriter
import pandas as pd
import numpy as np


def getsqldata(server, database, username, password, sql):
    conn = pymssql.connect(server, username, password, database, charset='utf8')
    cursor = conn.cursor()
    cursor.excute(sql)
    sqldata = cursor.fetchall()
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
                rsvdata[0].append(codeframe['date'][i])
                rsvdata[1].append(codeframe['code'][i])
                closeprice = codeframe['close'][i]
                maxhighprice = max(codeframe['high'][i - n: i])
                minlowprice = min(codeframe['low'][i - n: i])
                rsvdata[2].append(closeprice)
                rsvdata[3].append(maxhighprice)
                rsvdata[4].append(minlowprice)
                rsvdata[5].append((closeprice - minlowprice) / (maxhighprice - minlowprice) * 100)
                rsvdata[6].append(codeframe['avgprice'][i])
        else:
            continue
    # 计算K值, D值
    dataframe2 = pd.DataFrame(list(rsvdata), columns=['date', 'code', 'close', 'maxhigh', 'minlow', 'rsv', 'avgprice'])
    codelist2 = list(dataframe2.drop_duplicates('code')['code'].values)
    # kdvaluedata = [[date], [code], [return], [rsv], [kvalue], [dvalue]]
    kdvaluedata = [[], [], [], [], []]
    for jtem in codelist2:
        rsvframe = dataframe2[dataframe2.code == jtem]
        rsvframe = rsvframe.sort_values(by='date')
        for j in range(m1, len(rsvframe)):
            kdvaluedata[0].append(rsvframe['date'][j])
            kdvaluedata[1].append(rsvframe['code'][j])
            kdvaluedata[2].append(rsvframe['avgprice'][j])
            kdvaluedata[3].append(rsvframe['rsv'][j])
            kdvaluedata[4].append(np.average(rsvframe['rsv'][j - m1: j]))
            if j > m2:
                kdvaluedata[5].append(np.average(kdvaluedata[4][j - m2:j]))
            else:
                kdvaluedata[5].append(np.nan)
    kdvalueframe = pd.DataFrame(list(kdvaluedata), columns=['date', 'code', 'avgprice', 'rsv', 'kvalue', 'dvalue'])
    # 筛选条件: D值大于80
    kdframe = kdvalueframe[(np.isnan(kdvalueframe.dvalue) == False) & (kdvalueframe.dvalue > keepvalue)]

