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
    codereturn = sheet1.col_values(3, start_rowx=1)
    tdate = []
    for item in tradedate:
        if item not in tdate:
            tdate.append(item)
    for iDate in tdate:
        data1 = []
        for ii in range(0, len(tradedate)):
            if tradedate[ii] == iDate:
                data1.append([tradedate[ii], adj1[ii], codereturn[ii]])
        onelist = []
        data1.sort(key=lambda x: x[1])
        length = len(data1) / 20
        if length > 0:
            avg1 = sum([jj[2] for jj in data1[: length]]) / length - sum([jj[2] for jj in data1[-length:]]) / length
        else:
            avg1 = 0
        s11 = Series([float(jj[1]) for jj in data1])
        s21 = Series([float(jj[2]) for jj in data1])
        correl1 = s11.corr(s21)
        onelist.extend([float(avg1), 0 if isnan(correl1) else float(correl1)])
        result[iDate] = list(onelist + [len(data1)])
    return result


if __name__ == '__main__':
    today = datetime.datetime.now().strftime('%Y%m%d')
    filepath = 'C:\Users\Administrator\Desktop\work\/re_sql\/factor\LastStmNetInc.xlsx'
    result1 = operatesheet(0, filepath)
    dataList1 = [list([j] + result1[j]) for j in result1.keys()]
    dataList1.sort(key=lambda x: x[0])
    avgList1 = []
    for i in range(1, 3):
        avg11 = sum([j[i] for j in dataList1]) / len(dataList1)
        avgList1.append(avg11)
    field = [u'日期', u'收益_63', u'相关系数_63', u'家数_63']
    field3 = [u'63_平均收益', u'63_平均相关系数']
    wb1 = xlsxwriter.Workbook('C:\Users\Administrator\Desktop\LastStmNetInc_%s.xlsx' % today)
    sheet = wb1.add_worksheet('sheet1')
    # 中文字符格式
    allWordStyle = wb1.add_format({
        'bold': False,  # 字体加粗
        'border': 1,  # 添加边框
        'align': 'center',
        'valign': 'vcenter'
    })
    # 数值格式：添加边框，保留四位小数，负数显红色
    allNumStyle = wb1.add_format({
        'bold': False,  # 字体加粗
        'border': 1,  # 添加边框
        'align': 'center',
        'valign': 'vcenter',
        'num_format': '0.0000_ ;[red]-0.0000'
    })
    # 日期格式： yyyy/mm/dd
    dateStyle = wb1.add_format({
        'bold': False,  # 字体加粗
        'align': 'center',
        'valign': 'vcenter',
        'num_format': 'yyyy/m/d'
    })
    # 数值格式： 无边框，保留四位小数，负数显红色
    decimalStyle = wb1.add_format({
        'bold': False,  # 字体加粗
        'align': 'center',
        'valign': 'vcenter',
        'num_format': '0.0000_ ;[red]-0.0000'
    })
    # 单元格自动换行
    wrapStyle = wb1.add_format()
    wrapStyle.set_text_wrap()
    # 设置单元格格式
    sheet.set_column('A:A', 12, dateStyle)
    sheet.set_column('B:C', 8, decimalStyle)
    sheet.set_row(0, cell_format=wrapStyle)
    sheet.set_column('G:G', 20)
    sheet.set_column('H:H', 8)
    sheet.write_column('G2', field3, allWordStyle)
    sheet.write_column('H2', avgList1, allNumStyle)
    for i in range(0, len(field)):
        sheet.write(0, i, field[i])
    for row in range(1, len(dataList1) + 1):
        for col in range(0, len(field)):
            sheet.write(row, col, dataList1[row - 1][col])
    wb1.close()
    print 'Done!!!'
