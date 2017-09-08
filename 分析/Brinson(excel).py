# -*- coding:utf8 -*-
# Author: Youk Lin
# 业绩归因，计算组合每日的行业权重及收益
# Excel列名：1、日期，2、代码，3、市值，4、当日买入金额， 5、当日净买入金额
import xlrd
import xlsxwriter
import datetime
from numpy import *
import pandas as pd
from math import *


def operatesheet(num):
    sheet = wb.sheet_by_index(num)
    datelist = sheet.col_values(1, start_rowx=1)  # 交易日期
    codelist = sheet.col_values(6, start_rowx=1)  # 证券代码
    mvlist = sheet.col_values(8, start_rowx=1)  # 市值
    buylist = sheet.col_values(10, start_rowx=1)  # 当日买入金额
    netbuylist = sheet.col_values(12, start_rowx=1)  # 当日净买入金额
    sheetname = sheet.name
    tdate = []
    for item in datelist:
        if item not in tdate:
            tdate.append(item)
    tdate.sort()
    # 读取申万行业
    workbook = xlrd.open_workbook(r'..\..\excel\SW_Industry.xlsx')
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
    # 合并基金下不同组合同一股票持仓
    adjlist = dataframe1.groupby(['datelist', 'codelist', 'industrylist']).sum()
    # 计算每天总市值,{日期：当天总市值}
    summvlist = dataframe2.groupby(['datelist']).sum()
    summvdict = {}
    adjdict = {}
    for ktem in summvlist['mvlist'].keys():
        summvdict[ktem] = summvlist['mvlist'][ktem]
    # adjdict = {(date, code, industry): [mv, buy, netbuy]}
    for ntem in adjlist['mvlist'].keys():
        adjdict[ntem[0], ntem[1], ntem[2]] = [adjlist['mvlist'][ntem], adjlist['buylist'][ntem], adjlist['netbuylist'][ntem]]
    # 计算每只股票收益率和权重(取前一日的权重和当日的收益率)
    resultdict = {'date': [], 'industry': [], 'oneReturn': [], 'weight': []}
    for kk in range(1, len(tdate)):
        for jj in adjdict.keys():
            if jj[0] == tdate[kk] and adjdict.has_key((tdate[kk - 1], jj[1], jj[2])):
                resultdict['date'].append(jj[0])
                resultdict['industry'].append(jj[2])
                jl = (tdate[kk - 1], jj[1], jj[2])
                resultdict['oneReturn'].append((adjdict[jj][0] - adjdict[jl][0] - adjdict[jj][2]) / (adjdict[jl][0] + adjdict[jj][1]))
                resultdict['weight'].append(adjdict[jl][0] / summvdict[jl[0]])
    resultframe = pd.DataFrame(resultdict)
    resultframe['codeReturn'] = resultframe['oneReturn'] * resultframe['weight']
    # 计算行业权重及收益（前一日权重和 总收益率【未乘以权重之前的总收益】）
    adjresult = resultframe.groupby(['date', 'industry']).sum()
    adjresult['return'] = adjresult['codeReturn'] / adjresult['weight']
    result = []
    adjresultdict = {}
    for mtem in adjresult['return'].keys():
        result.append([mtem[0], mtem[1], adjresult['return'][mtem], adjresult['weight'][mtem]])
        adjresultdict[(mtem[0], mtem[1])] = [adjresult['return'][mtem], adjresult['weight'][mtem]]
    # 计算基金每日整体收益
    datereturn = resultframe['codeReturn'].groupby(resultframe['date']).sum()
    datereturnlist = [[datereturn.keys()[ll], datereturn[ll]] for ll in range(0, len(datereturn))]
    return result, sheetname, datereturnlist, adjresultdict


def readexcel(readfilepath, choosenum):
    workbook = xlrd.open_workbook(readfilepath)
    worksheet = workbook.sheet_by_index(0)
    valuelist0 = worksheet.col_values(0, start_rowx=1)
    valuelist1 = worksheet.col_values(1, start_rowx=1)
    valuelist2 = worksheet.col_values(2, start_rowx=1)

    if choosenum == 2:
        valuedict = {valuelist1[ii]: valuelist2[ii] for ii in range(0, len(valuelist1))}
    elif choosenum == 3:
        valuedict = {(valuelist0[ii], valuelist1[ii]): valuelist2[ii] for ii in range(0, len(valuelist1))}
    else:
        valuelist3 = worksheet.col_values(3, start_rowx=1)
        valuelist4 = worksheet.col_values(4, start_rowx=1)
        valuedict = [[valuelist0[ii], valuelist1[ii], valuelist2[ii], valuelist3[ii], valuelist4[ii]] for ii in range(0, len(valuelist1))]
    return valuedict

if __name__ == '__main__':
    sheetNum = input('input the num of sheet(larger than zero): ')
    today = datetime.datetime.now().strftime('%Y%m%d')
    localPath = '..\..\excel\\'
    filePath = localPath + 'chicang.xlsx'
    readFilePath1 = localPath + 'FundStandard.xlsx'
    readFilePath2 = localPath + 'IndexReturn.xlsx'
    readFilePath3 = localPath + 'IndexIndRe.xlsx'
    writewb1 = localPath + 'GroupAnalysis_%s.xlsx' % today
    writewb2 = localPath + 'AllDateReturn_%s.xlsx' % today
    writewb3 = localPath + 'Brinson_%s.xlsx' % today
    # 字典{Num: IndexCode}
    fundStandard = readexcel(readFilePath1, 2)
    # 字典{(date, IndexCode): IndexReturn}
    indexReturn = readexcel(readFilePath2, 3)
    # 列表[[date, IndexCode, Industry, Return, Weight]]
    indexIndRe = readexcel(readFilePath3, 5)
    indexIndRe.sort(key=lambda x: x[0])
    wb = xlrd.open_workbook(filePath)
    wb1 = xlsxwriter.Workbook(writewb1)
    wb2 = xlsxwriter.Workbook(writewb2)
    wb3 = xlsxwriter.Workbook(writewb3)
    field1 = [u'日期', u'行业', u'收益', u'权重', u'调整因子']
    field2 = [u'日期', u'基金收益', u'基准收益', u'调整因子', u'基金复合收益率', u'基准复合收益率']
    field3 = [u'日期', u'基准代码', u'基准行业', u'基准行业收益', u'基准行业权重', u'基金行业收益', u'基金行业权重', 'AR', 'SR', 'RAAK', 'RSSK']
    for k in range(0, sheetNum):
        # result1 = [[date, industry, return, weight]];
        # dayReturnList = [[date, fundReturn, indexReturn, kValue, RPK, RBK]]
        # resultDict1 = {(date, industry): [return, weight]}
        result1, sheetName, dayReturnList, resultDict1 = operatesheet(k)
        indexCode = fundStandard[k]
        dayReturnList.sort(key=lambda x: x[0])
        dayList = [dtem[0] for dtem in dayReturnList]
        dayList.sort()
        for j in range(0, len(dayReturnList)):
            indexDayReturn = indexReturn[(dayReturnList[j][0], indexCode)]
            fundReturn = float(dayReturnList[j][1])
            # 计算调整因子k值
            if indexDayReturn == fundReturn:
                kValue = 1.0 / (1 + fundReturn)
            else:
                kValue = (log(1 + fundReturn) - log(1 + indexDayReturn)) / (fundReturn - indexDayReturn)
            # 计算基金实际组合、基准组合的k期复合收益率
            if j == 0:
                RPK = fundReturn + 1
                RBK = indexDayReturn + 1
            else:
                # noinspection PyTypeChecker
                RPK = (dayReturnList[j - 1][4] + 1) * (fundReturn + 1)
                # noinspection PyTypeChecker
                RBK = (dayReturnList[j - 1][5] + 1) * (indexDayReturn + 1)
            dayReturnList[j].extend([indexDayReturn, kValue, RPK - 1, RBK - 1])
        kValueDict = {ktem[0]: ktem[3] for ktem in dayReturnList}
        for jtem in result1:
            jtem.append(kValueDict[jtem[0]])
        # 计算积极资产配置组合的k期复合收益率RAAK，积极股票选择组合的k期复合收益率RSSK
        # 以基准行业为准，计算得出以下字典fundIndexDict = {日期， 基准代码， 基准行业， 基准行业收益， 基准行业权重，基金行业收益， 基金行业权重， AR， SR, RAAK, RSSK}
        fundIndexDict = {'Date': [], 'IndexCode': [], 'Industry': [], 'IndexReturn': [], 'IndexWeight': [], 'FundReturn': [], 'FundWeight': [], 'AR': [], 'SR': []}
        for ltem in indexIndRe:
            if ltem[1] == indexCode and ltem[0] in dayList:
                fundIndexDict['Date'].append(ltem[0])
                fundIndexDict['IndexCode'].append(ltem[1])
                fundIndexDict['Industry'].append(ltem[2])
                fundIndexDict['IndexReturn'].append(ltem[3])
                fundIndexDict['IndexWeight'].append(ltem[4])
                if resultDict1.has_key((ltem[0], ltem[2])):
                    fundIndexDict['FundReturn'].append(resultDict1[(ltem[0], ltem[2])][0])
                    fundIndexDict['FundWeight'].append(resultDict1[(ltem[0], ltem[2])][1])
                    fundIndexDict['AR'].append((resultDict1[(ltem[0], ltem[2])][1] - ltem[4]) * ltem[3])
                    fundIndexDict['SR'].append((resultDict1[(ltem[0], ltem[2])][0] - ltem[3]) * ltem[4])
                else:
                    fundIndexDict['FundReturn'].append(0)
                    fundIndexDict['FundWeight'].append(0)
                    fundIndexDict['AR'].append(ltem[3] * ltem[4])
                    fundIndexDict['SR'].append(ltem[3] * ltem[4])
        # 计算积极资产配置组合RAAK和积极股票选择组合RSSK的k期复合收益率
        dayFundIndexFrame = pd.DataFrame(fundIndexDict)
        dayFundIndex = dayFundIndexFrame[['AR', 'SR']].groupby(dayFundIndexFrame['Date']).sum().sort_index()
        raaRssDict = {}
        fundIndexDict['RAAK'] = []
        fundIndexDict['RSSK'] = []
        for i in range(0, len(dayFundIndex.index)):
            raaRssDict[dayFundIndex.index[i]] = [reduce(lambda x, y: (x + 1) * (y + 1) - 1, dayFundIndex['AR'][: i + 1]),
                                                 reduce(lambda x, y: (x + 1) * (y + 1) - 1, dayFundIndex['SR'][: i + 1])]
        for atem in fundIndexDict['Date']:
            fundIndexDict['RAAK'].append(raaRssDict[atem][0])
            fundIndexDict['RSSK'].append(raaRssDict[atem][1])
        # 计算Q1\Q2\Q3\Q4
        fundIndexFrame = pd.DataFrame(fundIndexDict)
        # noinspection PyTypeChecker
        Q2 = sum(fundIndexFrame['FundWeight'] * fundIndexFrame['IndexReturn'] * (fundIndexFrame['RAAK'] + 1))
        # noinspection PyTypeChecker
        Q3 = sum(fundIndexFrame['IndexWeight'] * fundIndexFrame['FundReturn'] * (fundIndexFrame['RSSK'] + 1))
        Q4 = sum(map(lambda x: x[1] * (x[4] + 1), dayReturnList))
        Q1 = sum(map(lambda x: x[2] * (x[5] + 1), dayReturnList))
        fundIndexKeys = ['Date', 'IndexCode', 'Industry', 'IndexReturn', 'IndexWeight', 'FundReturn', 'FundWeight', 'AR', 'SR', 'RAAK', 'RSSK']
        sheet1 = wb1.add_worksheet(sheetName)
        sheet2 = wb2.add_worksheet(sheetName)
        sheet3 = wb3.add_worksheet(sheetName)
        sheet1.set_column('A:A', 11)
        sheet2.set_column('A:A', 11)
        sheet3.set_column('A:A', 11)
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
        for i in range(0, len(field3)):
            sheet3.write(0, i, field3[i])
        for col in range(0, len(fundIndexKeys)):
            sheet3.write_column(1, col, fundIndexDict[fundIndexKeys[col]])
        sheet3.write_column(1, 12, ['Q1', 'Q2', 'Q3', 'Q4'])
        sheet3.write_column(1, 13, [Q1, Q2, Q3, Q4])
    wb1.close()
    print 'Done!!!'
