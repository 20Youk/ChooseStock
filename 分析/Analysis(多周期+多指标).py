# -*- coding:utf8 -*-
# 分析选股指标与收益率的百分位收益差以及相关系数
# Excel列名：1、日期，2、代码，3、指标值，4、收益率
import xlrd
from pandas import *
import xlsxwriter
from math import *
import datetime


def operatesheet(num, path):
    result = {}
    wb = xlrd.open_workbook(path)
    sheet1 = wb.sheet_by_index(num)
    tradedate = sheet1.col_values(0, start_rowx=1)
    adj1 = sheet1.col_values(2, start_rowx=1)
    # adj2 = sheet1.col_values(3, start_rowx=1)
    # adj3 = sheet1.col_values(4, start_rowx=1)
    codereturn = sheet1.col_values(3, start_rowx=1)
    tdate = []
    for item in tradedate:
        if item not in tdate:
            tdate.append(item)
    for iDate in tdate:
        data1 = []
        for ii in range(0, len(tradedate)):
            if tradedate[ii] == iDate:
                # data1.append([tradedate[ii], adj1[ii], adj2[ii], adj3[ii], codereturn[ii]])
                data1.append([tradedate[ii], adj1[ii], codereturn[ii]])
        onelist = []
        # for kk in range(1, 4):
        for kk in range(1, 2):
            data1.sort(key=lambda x: x[kk], reverse=True)
            length = len(data1) / 20
            if length > 0:
                avg1 = sum([jj[2] for jj in data1[0: length]]) / length - sum([jj[2] for jj in data1[-length:]]) / length
                # avg1 = sum([jj[2] for jj in data1[-length:]]) / length
            else:
                avg1 = 0
            s11 = Series([float(jj[kk]) for jj in data1])
            s21 = Series([float(jj[2]) for jj in data1])
            correl1 = s11.corr(s21)
            onelist.extend([float(avg1), 0 if isnan(correl1) else float(correl1)])
        result[iDate] = list(onelist + [len(data1)])
        # result[iDate] = list(onelist + [len(data1) / 20])
    return result


if __name__ == '__main__':
    today = datetime.datetime.now().strftime('%Y%m%d')
    filepath = 'C:\Users\Administrator\Desktop\work\/re_sql\/factor\PctChg.xlsx'
    result1 = operatesheet(0, filepath)
    dataList1 = [list([j] + result1[j]) for j in result1.keys()]
    dataList1.sort(key=lambda x: x[0])
    result2 = operatesheet(1, filepath)
    dataList2 = [list([j] + result2[j]) for j in result2.keys()]
    dataList2.sort(key=lambda x: x[0])
    avgList1 = []
    avgList2 = []
    for i in range(1, 3):
        avg11 = sum([j[i] for j in dataList1]) / len(dataList1)
        avg21 = sum([j[i] for j in dataList2]) / len(dataList2)
        avgList1.append(avg11)
        avgList2.append(avg21)
    # field = [u'日期', u'MA5收益差_21', u'MA5相关系数_21', u'MA10收益差_21', u'MA10相关系数_21', u'MA20收益差_21', u'MA20相关系数_21', u'家数_21']
    # field2 = [u'日期', u'MA5收益差_63', u'MA5相关系数_63', u'MA10收益差_63', u'MA10相关系数_63', u'MA20收益差_63', u'MA20相关系数_63',u'家数_63']
    # field3 = [u'21_MA5平均收益差', u'21_MA5平均相关系数', u'21_MA10平均收益差', u'21_MA10平均相关系数', u'21_MA20平均收益差', u'21_MA20平均相关系数',
    #           u'63_MA5平均收益差', u'63_MA5平均相关系数', u'63_MA10平均收益差', u'63_MA10平均相关系数', u'63_MA20平均收益差', u'63_MA20平均相关系数']
    field = [u'日期', u'收益_21', u'相关系数_21', u'家数_21']
    field2 = [u'日期', u'收益_63', u'相关系数_63', u'家数_63']
    field3 = [u'21_平均收益', u'21_平均相关系数', u'63_平均收益', u'63_平均相关系数']
    wb1 = xlsxwriter.Workbook('C:\Users\Administrator\Desktop\PctChgAnalysis_%s.xlsx' % today)
    sheet = wb1.add_worksheet('sheet1')
    allWordStyle = wb1.add_format({
        'bold': False,  # 字体加粗
        'border': 1,  # 添加边框
        'align': 'center',
        'valign': 'vcenter'
    })
    allNumStyle = wb1.add_format({
        'bold': False,  # 字体加粗
        'border': 1,  # 添加边框
        'align': 'center',
        'valign': 'vcenter',
        'num_format': '0.0000_ ;[red]-0.0000'
    })
    dateStyle = wb1.add_format({
        'bold': False,  # 字体加粗
        'align': 'center',
        'valign': 'vcenter',
        'num_format': 'yyyy/m/d'
    })
    decimalStyle = wb1.add_format({
        'bold': False,  # 字体加粗
        'align': 'center',
        'valign': 'vcenter',
        'num_format': '0.0000_ ;[red]-0.0000'
    })
    wrapStyle = wb1.add_format()
    wrapStyle.set_text_wrap()
    sheet.set_column('A:A', 12, dateStyle)
    sheet.set_column('B:C', 8, decimalStyle)
    sheet.set_column('J:J', 12, dateStyle)
    sheet.set_column('K:L', 8, decimalStyle)
    sheet.set_row(0, cell_format=wrapStyle)
    sheet.set_column('S:S', 20)
    sheet.set_column('T:T', 8)
    sheet.write_column('S2', field3, allWordStyle)
    sheet.write_column('T2', avgList1 + avgList2, allNumStyle)
    for i in range(0, len(field)):
        sheet.write(0, i, field[i])
    for row in range(1, len(dataList1) + 1):
        for col in range(0, len(field)):
            sheet.write(row, col, dataList1[row - 1][col])
    for i in range(0, len(field2)):
        sheet.write(0, i + 9, field2[i])
    for row in range(1, len(dataList2) + 1):
        for col in range(0, len(field2)):
            sheet.write(row, col + 9, dataList2[row - 1][col])
    wb1.close()
    print 'Done!!!'
