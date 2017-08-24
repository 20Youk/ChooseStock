# -*- coding:utf8 -*-
import xlrd
from pandas import *
import xlsxwriter
from math import *
import datetime
today = datetime.datetime.now().strftime('%Y%m%d')
wb = xlrd.open_workbook('C:\Users\Administrator\Desktop\QualityAnalyse_20170809.xlsx')
result1 = {}
result2 = {}
result3 = {}
result4 = {}
result0 = {}
sheet1 = wb.sheet_by_index(0)
sheet2 = wb.sheet_by_index(1)
sheet3 = wb.sheet_by_index(2)
sheet4 = wb.sheet_by_index(3)
tradeDate1 = sheet1.col_values(0, start_rowx=1)
tradeDate2 = sheet2.col_values(0, start_rowx=1)
tradeDate3 = sheet3.col_values(0, start_rowx=1)
tradeDate4 = sheet4.col_values(0, start_rowx=1)
tradeAssetsAdj = sheet1.col_values(2, start_rowx=1)
codeReturn1 = sheet1.col_values(3, start_rowx=1)
liabAdj = sheet2.col_values(2, start_rowx=1)
codeReturn2 = sheet2.col_values(3, start_rowx=1)
roeAdj = sheet3.col_values(2, start_rowx=1)
codeReturn3 = sheet3.col_values(3, start_rowx=1)
cashAdj = sheet4.col_values(2, start_rowx=1)
codeReturn4 = sheet4.col_values(3, start_rowx=1)
code1 = sheet1.col_values(1, start_rowx=1)
code2 = sheet2.col_values(1, start_rowx=1)
code3 = sheet3.col_values(1, start_rowx=1)
code4 = sheet4.col_values(1, start_rowx=1)
dict0 = {'MinDate': list(tradeDate1 + tradeDate2 + tradeDate3 + tradeDate4), 'Code': list(code1 + code2 + code3 + code4), 'Adj': list(tradeAssetsAdj + liabAdj + roeAdj + cashAdj), 'CodeReturn': list(codeReturn1 + codeReturn2 + codeReturn3 + codeReturn4)}
dateFrame0 = DataFrame(dict0)
dataDict0 = dateFrame0.groupby(['Code', 'MinDate']).mean()
tradeDate0 = [j[1] for j in dataDict0['Adj'].keys()]
Adj0 = list(dataDict0['Adj'])
codeReturn0 = list(dataDict0['CodeReturn'])
tDate = []
for item in tradeDate1:
    if item not in tDate:
        tDate.append(item)
for iDate in tDate:
    data1 = []
    data2 = []
    data3 = []
    data4 = []
    data0 = []
    for i in range(0, len(tradeDate1)):
        if tradeDate1[i] == iDate:
            data1.append([tradeDate1[i], tradeAssetsAdj[i], codeReturn1[i]])
    for i in range(0, len(tradeDate2)):
        if tradeDate2[i] == iDate:
            data2.append([tradeDate2[i], liabAdj[i], codeReturn2[i]])
    for i in range(0, len(tradeDate3)):
        if tradeDate3[i] == iDate:
            data3.append([tradeDate3[i], roeAdj[i], codeReturn3[i]])
    for i in range(0, len(tradeDate4)):
        if tradeDate4[i] == iDate:
            data4.append([tradeDate4[i], cashAdj[i], codeReturn4[i]])
    for i in range(0, len(tradeDate0)):
        if tradeDate0[i] == iDate:
            data0.append([tradeDate0[i], Adj0[i], codeReturn0[i]])
    data1.sort(key=lambda x: x[1], reverse=True)
    data2.sort(key=lambda x: x[1], reverse=True)
    data3.sort(key=lambda x: x[1], reverse=True)
    data4.sort(key=lambda x: x[1], reverse=True)
    data0.sort(key=lambda x: x[1], reverse=True)
    length = int(ceil(len(data1) / 20))
    length2 = int(ceil(len(data2) / 20))
    length3 = int(ceil(len(data3) / 20))
    length4 = int(ceil(len(data4) / 20))
    length0 = int(ceil(len(data0) / 20))
    avg0 = sum([j[2] for j in data0[0: length0]]) / length0 - sum([j[2] for j in data0[-length0:]]) / length0 if length > 0 else 0
    s10 = Series([j[1] for j in data0])
    s20 = Series([j[2] for j in data0])
    correl0 = s10.corr(s20)
    avg1 = sum([j[2] for j in data1[0: length]]) / length - sum([j[2] for j in data1[-length:]]) / length if length > 0 else 0
    s11 = Series([j[1] for j in data1])
    s21 = Series([j[2] for j in data1])
    correl1 = s11.corr(s21)
    avg2 = sum([j[2] for j in data2[0: length2]]) / length2 - sum([j[2] for j in data2[-length2:]]) / length2 if length2 > 0 else 0
    s12 = Series([j[1] for j in data2])
    s22 = Series([j[2] for j in data2])
    correl2 = s12.corr(s22)
    avg3 = sum([j[2] for j in data3[0: length3]]) / length3 - sum([j[2] for j in data3[-length3:]]) / length3 if length3 > 0 else 0
    s13 = Series([j[1] for j in data3])
    s23 = Series([j[2] for j in data3])
    correl3 = s13.corr(s23)
    avg4 = sum([j[2] for j in data4[0: length4]]) / length4 - sum([j[2] for j in data4[-length4:]]) / length4 if length4 > 0 else 0
    s14 = Series([j[1] for j in data4])
    s24 = Series([j[2] for j in data4])
    correl4 = s14.corr(s24)
    result1[iDate] = [float(avg1), 0 if isnan(correl1) else float(correl1), len(data1)]
    result2[iDate] = [float(avg2), 0 if isnan(correl2) else float(correl2), len(data2)]
    result3[iDate] = [float(avg3), 0 if isnan(correl3) else float(correl3), len(data3)]
    result4[iDate] = [float(avg4), 0 if isnan(correl4) else float(correl4), len(data4)]
    result0[iDate] = [float(avg0), 0 if isnan(correl0) else float(correl0)]
dataList = [list([j] + result1[j] + result2[j] + result3[j] + result4[j] + result0[j]) for j in result1.keys()]
dataList.sort(key=lambda x: x[0])
field = [u'日期', u'可投资资本收益差', u'可投资资本相关系数', u'可投资资本家数', u'负债指标收益差', u'负债指标相关系数', u'负债指标家数', u'ROE收益差', u'ROE相关系数', u'ROE指标家数', u'净现金收益差', u'净现金相关系数', u'净现金家数', u'综合收益差', u'综合相关系数']
wb1 = xlsxwriter.Workbook('C:\Users\Administrator\Desktop\QualityAnalysis_%s.xlsx' % today)
sheet = wb1.add_worksheet('sheet1')
sheet.set_column('A:A', 11)
for i in range(0, len(field)):
    sheet.write(0, i, field[i])
for row in range(1, len(dataList) + 1):
    for col in range(0, len(field)):
        sheet.write(row, col, dataList[row - 1][col])
wb1.close()
print 'Done!!!'
