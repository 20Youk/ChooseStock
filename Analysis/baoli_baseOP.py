# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用: 保理系统基础数据分析:融资日期分布
import xlrd
import xlsxwriter
import datetime

wb = xlrd.open_workbook('../../file/piaoju_base.xlsx')
ws = wb.sheet_by_index(1)
supplier_list = ws.col_values(1, start_rowx=1)
hotel_list = ws.col_values(2, start_rowx=1)
startDate_list = ws.col_values(3, start_rowx=1)
endDate_list = ws.col_values(4, start_rowx=1)
amount_list = ws.col_values(5, start_rowx=1)
supplier = supplier_list[0]
hotel = hotel_list[0]
startDate = startDate_list[0]
endDate = endDate_list[0]
result_data = []
one_list = [[supplier_list[0], hotel_list[0], startDate, 1],
            [supplier_list[0], hotel_list[0], endDate, 1]]
firstDay = datetime.datetime(1900, 1, 1)
today = datetime.datetime.now()
today_int = (today - firstDay).days + 2
for i in range(1, len(supplier_list)):
    if supplier_list[i] == supplier and hotel_list[i] == hotel:
        if startDate <= startDate_list[i] <= endDate <= endDate_list[i]:
            endDate = endDate_list[i]
        elif endDate < startDate_list[i] - 1:
            result_data.append([supplier_list[i], hotel_list[i], endDate + 1, startDate_list[i] - 1])
            startDate = startDate_list[i]
            endDate = endDate_list[i]
        elif endDate == startDate_list[i]:
            endDate = endDate_list[i]
        elif endDate == startDate_list[i] - 1:
            endDate = endDate_list[i]
    else:
        result_data.append([supplier, hotel, endDate + 1, today_int])
        supplier = supplier_list[i]
        hotel = hotel_list[i]
        startDate = startDate_list[i]
        endDate = endDate_list[i]
        continue
wb_1 = xlsxwriter.Workbook('../../file/baoli_date.xlsx')
ws_1 = wb_1.add_worksheet(u'供应商未融资日期分布')
dateType = wb_1.add_format()
dateType.set_num_format('yyyy-mm-dd')
field = [u'供应商', u'酒店', u'未融资开始日期', u'未融资结束日期']
ws_1.write_row(0, 0, field)
maxLen_1 = len(result_data[0][0])
maxLen_2 = len(result_data[0][1])
for j in range(0, len(result_data)):
    maxLen_1 = max(maxLen_1, len(result_data[j][0]))
    maxLen_2 = max(maxLen_2, len(result_data[j][1]))
    ws_1.write_row(j+1, 0, result_data[j][:2])
    ws_1.write_row(j+1, 2, result_data[j][2:], dateType)
ws_1.set_column('A:A', maxLen_1)
ws_1.set_column('B:B', maxLen_2)
ws_1.set_column('C:D', 14)
wb_1.close()
print 'Done!!!'
