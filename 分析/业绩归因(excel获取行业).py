# -*- coding:utf8 -*-
# 业绩归因，计算组合每日的行业权重及收益
# Excel列名：1、日期，2、代码，3、市值，4、当日买入金额， 5、当日净买入金额
import xlrd
import xlsxwriter
import datetime
from numpy import *
import pandas as pd
from math import *


def operatesheet(num):
    sheet1 = wb.sheet_by_index(num)
    datelist = sheet1.col_values(1, start_rowx=1)  # 交易日期
    codelist = sheet1.col_values(6, start_rowx=1)  # 证券代码
    mvlist = sheet1.col_values(8, start_rowx=1)  # 市值
    buylist = sheet1.col_values(10, start_rowx=1)  # 当日买入金额
    netbuylist = sheet1.col_values(12, start_rowx=1)  # 当日净买入金额
    sheetname = sheet1.name
    tdate = []
    for item in datelist:
        if item not in tdate:
            tdate.append(item)
    tdate.sort()
    # 读取申万行业
    workbook = xlrd.open_workbook(r'..\excel\SW_Industry.xlsx')
    worksheet = workbook.sheet_by_index(0)
    seccode = worksheet.col_values(0, start_rowx=1)
    swindustry = worksheet.col_values(1, start_rowx=1)
    industrydict = {seccode[ij]: swindustry[ij] for ij in range(0, len(seccode))}
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
    # 计算行业权重及收益
    adjresult = resultframe.groupby(['date', 'industry']).sum()
    result = []
    for mtem in adjresult['return'].keys():
        result.append([mtem[0], mtem[1], adjresult['return'][mtem], adjresult['weight'][mtem]])
    # 计算基金每日整体收益
    resultframe['weightreturn'] = resultframe['return'] * resultframe['weight']
    datereturn = resultframe['weightreturn'].groupby(resultframe['date']).sum()
    datereturnlist = [[datereturn.keys()[ll], datereturn[ll]] for ll in range(0, len(datereturn))]
    return result, sheetname, datereturnlist


def readexcel(readfilepath, choosenum):
    workbook = xlrd.open_workbook(readfilepath)
    worksheet = workbook.sheet_by_index(0)
    valuelist0 = worksheet.col_values(0, start_rowx=1)
    valuelist1 = worksheet.col_values(1, start_rowx=1)
    valuelist2 = worksheet.col_values(2, start_rowx=1)
    if choosenum == 2:
        valuedict = {valuelist1[ii]: valuelist2[ii] for ii in range(0, len(valuelist1))}
    else:
        valuedict = {(valuelist0[ii], valuelist1[ii]): valuelist2[ii] for ii in range(0, len(valuelist1))}
    return valuedict

if __name__ == '__main__':
    sheetNum = input('input the num of sheet(larger than zero): ')
    today = datetime.datetime.now().strftime('%Y%m%d')
    filePath = '..\excel\chicang.xlsx'
    readFilePath1 = '..\excel\FundStandard.xlsx'
    readFilePath2 = '..\excel\IndexReturn.xlsx'
    writewb1 = '..\excel\GroupAnalysis_%s.xlsx' % today
    writewb2 = '..\excel\AllDateReturn_%s.xlsx' % today
    fundStandard = readexcel(readFilePath1, 2)
    indexReturn = readexcel(readFilePath2, 3)
    wb = xlrd.open_workbook(filePath)
    wb1 = xlsxwriter.Workbook(writewb1)
    wb2 = xlsxwriter.Workbook(writewb2)
    field1 = [u'日期', u'行业', u'收益', u'权重', u'调整因子']
    field2 = [u'日期', u'基金收益', u'基准收益', u'调整因子']
    for k in range(0, sheetNum):
        result1, sheetName, dayReturnList = operatesheet(k)
        indexCode = fundStandard[k]
        for j in range(0, len(dayReturnList)):
            indexDayReturn = indexReturn[(dayReturnList[j][0], indexCode)]
            fundReturn = float(dayReturnList[j][1])
            kValue = (log(1 + fundReturn) - log(1 + indexDayReturn)) / (fundReturn - indexDayReturn)
            dayReturnList[j].extend([indexDayReturn, kValue])
        kValueDict = {ktem[0]: ktem[3] for ktem in dayReturnList}
        for jtem in result1:
            jtem.append(kValueDict[jtem[0]])
        sheet1 = wb1.add_worksheet(sheetName)
        sheet2 = wb2.add_worksheet(sheetName)
        sheet1.set_column('A:A', 11)
        sheet2.set_column('A:A', 11)
        for i in range(0, len(field1)):
            sheet1.write(0, i, field1[i])
        for row in range(1, len(result1) + 1):
            for col in range(0, len(field1)):
                sheet1.write(row, col, result1[row - 1][col])
        for i in range(0, len(field2)):
            sheet2.write(0, i, field2[i])
        for row in range(1, len(dayReturnList) + 1):
            for col in range(0, len(field2)):
                sheet2.write(row, col, dayReturnList[row - 1][col])
    wb1.close()
    print 'Done!!!'

