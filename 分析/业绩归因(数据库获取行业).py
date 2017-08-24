# -*- coding:utf8 -*-
# 业绩归因，计算组合每日的行业权重及收益
# Excel列名：1、日期，2、代码，3、市值，4、当日买入金额， 5、当日净买入金额
import xlrd
import xlsxwriter
import datetime
from numpy import *
import pandas as pd
import pymssql


def operatesheet(num):
    server = '192.168.8.200'
    user = 'sa'
    password = 'wind123@pa'
    database = 'PortfolioData'
    conn = pymssql.connect(server, user, password, database, charset='utf8')
    cursor = conn.cursor()
    sheet1 = wb.sheet_by_index(num)
    datelist = sheet1.col_values(1, start_rowx=1)  # 交易日期
    codelist = sheet1.col_values(6, start_rowx=1)  # 证券代码
    mvlist = sheet1.col_values(8, start_rowx=1)  # 市值
    buylist = sheet1.col_values(10, start_rowx=1)  # 当日买入金额
    netbuylist = sheet1.col_values(12, start_rowx=1)  # 当日净买入金额
    sheetname = sheet1.name
    tdate = []
    tcode = []
    for item in datelist:
        if item not in tdate:
            tdate.append(item)
    tdate.sort()
    for item in codelist:
        if item not in tcode:
            tcode.append(item)
    codestring = ('%s,' * len(tcode))[:-1]
    sql11 = sql1 % codestring
    cursor.execute(sql11, tuple(tcode))
    sqldata = cursor.fetchall()
    industrydict = dict(sqldata)
    industrylist = []
    for ii in range(0, len(codelist)):
        industrylist.append(industrydict[codelist[ii]])
    dict1 = {'datelist': datelist, 'codelist': codelist, 'industrylist': industrylist, 'mvlist': mvlist, 'buylist': buylist, 'netbuylist': netbuylist}
    dict2 = {'datelist': datelist, 'mvlist': mvlist}
    dataframe1 = pd.DataFrame(dict1, dtype=float)
    dataframe2 = pd.DataFrame(dict2, dtype=float)
    adjlist = dataframe1.groupby(['datelist', 'codelist', 'industrylist']).sum()
    summvlist = dataframe2.groupby(['datelist']).sum()
    summvdict = {}
    adjdict = {}
    for ktem in summvlist['mvlist'].keys():
        summvdict[ktem] = summvlist['mvlist'][ktem]
    for ntem in adjlist['mvlist'].keys():
        adjdict[ntem[0], ntem[1], ntem[2]] = [adjlist['mvlist'][ntem], adjlist['buylist'][ntem], adjlist['netbuylist'][ntem]]
    resultdict = {'date': [], 'industry': [], 'return': [], 'weight': []}
    for kk in range(1, len(tdate)):
        for jj in adjdict.keys():
            if jj[0] == tdate[kk] and adjdict.has_key((tdate[kk - 1], jj[1], jj[2])):
                resultdict['date'].append(jj[0])
                resultdict['industry'].append(jj[2])
                jl = [tdate[kk - 1], jj[1], jj[2]]
                resultdict['return'].append((adjdict[jj][0] - adjdict[tuple(jl)][0] - adjdict[jj][2]) / (adjdict[tuple(jl)][0] + adjdict[jj][1]))
                resultdict['weight'].append(adjdict[jj][0] / summvdict[jj[0]])
    resultframe = pd.DataFrame(resultdict)
    # 分行业计算权重及收益
    adjresult = resultframe.groupby(['date', 'industry']).sum()
    result = []
    for mtem in adjresult['return'].keys():
        result.append([mtem[0], mtem[1], adjresult['return'][mtem], adjresult['weight'][mtem]])
    # 分日期计算整体收益
    resultframe['weightreturn'] = resultframe['return'] * resultframe['weight']
    datereturn = resultframe['weightreturn'].groupby(resultframe['date']).sum()
    datereturnlist = [[datereturn.keys()[ll], datereturn[ll]] for ll in range(0, len(datereturn))]
    return result, sheetname, datereturnlist

if __name__ == '__main__':
    today = datetime.datetime.now().strftime('%Y%m%d')
    filepath = 'C:\Users\Administrator\Desktop\chicang.xlsx'
    sql1 = '''select SecCode, IndustryName1 from StockIndustry where SecCode in (%s) and EndDate is NULL and IndustryType = 1 '''
    wb = xlrd.open_workbook(filepath)
    field = [u'日期', u'行业', u'收益', u'权重']
    wb1 = xlsxwriter.Workbook('C:\Users\Administrator\Desktop\GroupAnalysis_%s.xlsx' % today)
    for k in range(0, 4):
        result1, sheetName, dateReturnList = operatesheet(k)
        sheet = wb1.add_worksheet(sheetName)
        sheet.set_column('A:A', 11)
        for i in range(0, len(field)):
            sheet.write(0, i, field[i])
        for row in range(1, len(result1) + 1):
            for col in range(0, len(field)):
                sheet.write(row, col, result1[row - 1][col])
    wb1.close()
    print 'Done!!!'
