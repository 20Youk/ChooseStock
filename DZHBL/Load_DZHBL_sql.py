# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:导出DZHBL数据库数据到Excel
import pymssql
import xlsxwriter
import time


def read_sql(host, user, password, database, sql):
    conn = pymssql.connect(host, user, password, database, charset='utf8')
    cursor = conn.cursor()
    cursor.execute(sql)
    data = cursor.fetchall()
    field = cursor.description
    conn.close()
    return data, field


def write_to_excel(filepath, data, field):
    wb = xlsxwriter.Workbook(filepath)
    ws = wb.add_worksheet('Sheet1')
    for i in range(0, len(field)):
        ws.write(0, i, field[i][0])
    for i in range(0, len(data)):
        for j in range(0, len(field)):
            ws.write(i + 1, j, data[i][j])
    wb.close()
    return

if __name__ == '__main__':
    excelPath = '../../file/Load_Data_%s.xlsx' % time.strftime('%Y%m%d')
    server = 'localhost'
    userName = '***'
    passWord = '*****'
    dataBase = '***'
    sql1 = '''with t2 as (select t1.DocEntry, convert(decimal(18,2), sum(t1.LineTotal)) as SumLineTotle,
                            convert(decimal(18,2), sum(t1.ALineTotal)) as SumALineTotle
                            from U_RZQR1 t1 group by t1.DocEntry)
            select t1.RZHCode as '融资申请编号',
                    t.CardName as '融资方名称',
                    t.BCardName as '买方酒店名称',
                    case when t1.RzType = 1 then '票前融资' else '票后融资' end as 融资方式,
                    year(t1.RMoth)+MONTH(t1.RMoth)*0.01 as 融资月份,
                    convert(date, t.SDate, 120) as '账期开始日期',
                    convert(date, t.EDate, 120) as '帐期结束日期',
                    t2.SumLineTotle as 应收款总额,
                    t2.SumALineTotle as 应放款总额,
                    convert(date, t.DocDate, 120) as '放款日期'
             from U_OPFK1 t, U_RZQR t1 , t2, U_OPFK t3
             where t.DocEntry = t3.DocEntry
                 and t3.DocType = 'F'
                 and t.QCode = t1.HCode
                 and t1.DocType = 'Q'
                 and t1.DocEntry = t2.DocEntry
                 and convert(date, t.DocDate, 120) >= convert(date, dateadd(mm, -1, dateadd(dd, -day(getdate())+1, GETDATE())), 120)
                 and convert(date, t.DocDate, 120) < convert(date, dateadd(dd, -day(getdate())+1, GETDATE()), 120)
                 order by convert(date, t.DocDate, 120) '''
    sqlData, field1 = read_sql(server, userName, passWord, dataBase, sql1)
    write_to_excel(excelPath, sqlData, field1)
    print 'Done!'
