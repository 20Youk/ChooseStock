# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:导出DZHBL数据库的融资时效基础数据到Excel
import pymssql
import xlsxwriter
import time
import os
import ConfigParser

con = ConfigParser.ConfigParser()
config_path = './config/config.txt'
with open(config_path, 'r') as f:
    con.readfp(f)
    server = con.get('info', 'server')
    userName = con.get('info', 'username')
    passWord = con.get('info', 'password')
    dataBase = con.get('info', 'database')


def read_sql(host, user, password, database, sql1, sql2, sql3):
    conn = pymssql.connect(host, user, password, database, charset='utf8')
    cursor = conn.cursor()
    cursor.execute(sql1)
    data1 = cursor.fetchall()
    field1 = cursor.description
    cursor.execute(sql2)
    data2 = cursor.fetchall()
    field2 = cursor.description
    cursor.execute(sql3)
    data3 = cursor.fetchall()
    field3 = cursor.description
    conn.close()
    return data1, field1, data2, field2, data3, field3


def write_to_excel(filepath, data1, field1, data2, field2, data3, field3):
    wb = xlsxwriter.Workbook(filepath)
    ws1 = wb.add_worksheet(u'运营融资时效')
    ws2 = wb.add_worksheet(u'供应商融资时效')
    ws3 = wb.add_worksheet(u'台账')
    newstyle = wb.add_format()
    newstyle.set_border(2)
    newstyle.set_font_size(9)
    newstyle.set_font_name(u'宋体')
    newstyle.set_align('left')      # 左对齐
    newstyle.set_align('vcenter')   # 垂直居中
    for i in range(0, len(field1)):
        ws1.write(0, i, field1[i][0], newstyle)
    for i in range(0, len(data1)):
        for j in range(0, len(field1)):
            ws1.write(i + 1, j, data1[i][j], newstyle)
    for i in range(0, len(field2)):
        ws2.write(0, i, field2[i][0], newstyle)
    for i in range(0, len(data2)):
        for j in range(0, len(field2)):
            ws2.write(i + 1, j, data2[i][j], newstyle)
    for i in range(0, len(field3)):
        ws3.write(0, i, field3[i][0], newstyle)
    for i in range(0, len(data3)):
        for j in range(0, len(field3)):
            ws3.write(i + 1, j, data3[i][j], newstyle)
    wb.close()
    return

if __name__ == '__main__':
    month = int(time.strftime('%m'))
    today = time.strftime('%Y%m%d')
    filePath = '/home/vftpuser/public/融资时效基础数据%s' % today
    if not os.path.exists(filePath):
        os.mkdir(filePath)
    excel = filePath + '/融资时效基础数据%s.xlsx' % today
    sql_1 = '''exec opration_aging'''
    sql_2 = '''exec supplier_aging'''
    sql_3 = '''exec taizhang'''
    sqlData_1, field_1, sqldata_2, field_2, sqldata_3, field_3 = read_sql(server, userName, passWord, dataBase, sql_1, sql_2, sql_3)
    write_to_excel(excel, sqlData_1, field_1, sqldata_2, field_2, sqldata_3, field_3)
    print 'Done!'
