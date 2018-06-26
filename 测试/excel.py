# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:
import xlsxwriter
import pandas as pd

df = pd.read_excel('../../file/hotel.xlsx', sheet_name='Sheet2')
field = [u'酒店', u'时间', u'入住率', u'宴会收入']
data = []
for i in range(0, len(df), 2):
    for j in range(2, len(df.columns)):
        data.append([df.columns[j], df['Month'][i], df[df.columns[j]][[i]], df[df.columns[j]][i + 1]])
data.sort()
wb = xlsxwriter.Workbook('../../file/data_result.xlsx')
ws = wb.add_worksheet('Sheet1')
ws.write_row(0, 0, field)
for k in range(0, len(data)):
    ws.write_row(k + 1, 0, data[k])
date_style = wb.add_format()
date_style.set_num_format('YYYY-MM-DD')
ws.set_column('B:B'
              '', 11, date_style)
wb.close()
