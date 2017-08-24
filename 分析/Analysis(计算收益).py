# -*- coding:utf8 -*-
# 分析选股指标与收益率的百分位收益差以及相关系数
# Excel列名：1、日期，2、代码，3、指标值，4、收益率
import xlrd
import xlsxwriter
import datetime
from numpy import *


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
        data1.sort(key=lambda x: x[1])
        length = len(data1) / 20
        if len(data1) > 0:
            avg = sum([jj[2] for jj in data1[: length]]) / length - 0.03
        else:
            avg = 0
        result[iDate] = [length, avg]
    return result


# 两个指标
def factors(num, path):
    result = {}
    wb = xlrd.open_workbook(path)
    sheet1 = wb.sheet_by_index(num)
    tradedate = sheet1.col_values(0, start_rowx=1)
    adj1 = sheet1.col_values(2, start_rowx=1)
    adj2 = sheet1.col_values(3, start_rowx=1)
    codereturn = sheet1.col_values(4, start_rowx=1)
    tdate = []
    for item in tradedate:
        if item not in tdate:
            tdate.append(item)
    for iDate in tdate:
        data1 = []
        for ii in range(0, len(tradedate)):
            if tradedate[ii] == iDate and adj2[ii] > 0.15:
                data1.append([tradedate[ii], adj1[ii], adj2[ii], codereturn[ii]])
        data1.sort(key=lambda x: x[1])
        if len(data1) > 0:
            avg = sum([jj[3] for jj in data1[: 300]]) / 300 - 0.03
        else:
            avg = 0
        result[iDate] = [len(data1), avg]
    return result


if __name__ == '__main__':
    today = datetime.datetime.now().strftime('%Y%m%d')
    filepath = 'C:\Users\Administrator\Desktop\work\/re_sql\/factor\PctChg.xlsx'
    result1 = operatesheet(1, filepath)
    # dataList1 = [日期， 家数， 收益率]
    dataList1 = [list([j] + result1[j]) for j in result1.keys()]
    dataList1.sort(key=lambda x: x[0])
    dayReturn = [j[2] for j in dataList1]
    avgReturn = sum(dayReturn) / len(dayReturn)
    riskProportion = std(dayReturn, ddof=1)
    riskReturn = avgReturn / riskProportion * sqrt(4)
    yearReturn = avgReturn * 240
    avgList1 = [avgReturn, riskProportion, riskReturn, yearReturn]
    dataList1[0].append(dataList1[0][2] + 1)
    for i in range(1, len(dataList1)):
        dataList1[i].append(dataList1[i - 1][3] * (1 + dataList1[i][2]))
    field = [u'日期', u'大于15%家数', u'收益_63', u'单位净值']
    field3 = [u'63_平均收益', u'63_风险率', u'63_收益风险比', u'年化收益率']
    wb1 = xlsxwriter.Workbook('C:\Users\Administrator\Desktop\PctchgReturn_%s.xlsx' % today)
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
    sheet.set_column('C:D', 8, decimalStyle)
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
    chart = wb1.add_chart({'type': 'line'})
    chart.add_series({
        'name': ['sheet1', 0, 3],
        'categories': ['sheet1', 1, 0, len(dataList1), 0],
        'values': ['sheet1', 1, 3, len(dataList1), 3],
        'line': {'colorIndex': '23'},
    })
    chart.set_title({'name': u'单位净值'})
    # chart.set_x_axis({'name': u'日期'})
    # chart.set_y_axis({'name': u'单位净值'})
    chart.set_size({'width': 500, 'height': 350})
    sheet.insert_chart('F7', chart)
    wb1.close()
    print 'Done!!!'
