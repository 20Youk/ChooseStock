# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:导出DZHBL数据库的账期逾期数据到Excel
import pymssql
import xlsxwriter
import time
import os


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
    newstyle = wb.add_format()
    newstyle.set_border(2)
    newstyle.set_font_size(9)
    newstyle.set_font_name(u'宋体')
    newstyle.set_align('left')      # 左对齐
    newstyle.set_align('vcenter')   # 垂直居中
    length1 = 10
    length2 = 10
    for i in range(0, len(field)):
        ws.write(0, i, field[i][0], newstyle)
    for i in range(0, len(data)):
        for j in range(0, len(field)):
            ws.write(i + 1, j, data[i][j], newstyle)
        length1 = max(length1, len(data[i][0]))
        length2 = max(length2, len(data[i][1]))
    ws.set_column('A:A', length1)
    ws.set_column('B:B', length2)
    wb.close()
    return

if __name__ == '__main__':
    month = int(time.strftime('%m'))
    today = time.strftime('%Y%m%d')
    # filePath = '../../file/%s' % today
    filePath = '/home/vftpuser/public/关注逾期数据/%s' % today
    if not os.path.exists(filePath):
        os.mkdir(filePath)
    excel = filePath + '/关注逾期数据%s.xlsx' % today
    server = '***'
    userName = '***'
    passWord = '***'
    dataBase = '***'
    sql1 = '''with t00 as (select t1.DocEntry, t2.BaseEntry,t2.QCode, t2.BDate, t2.BTotal from U_OPFK t1, U_OPFK1 t2 where t1.DocEntry = t2.DocEntry and  t2.BaseEntry is not null and t1.DocType = 'T')
            ,a as (select
            t.BCardName,
             t.CardName,
             datediff(dd, convert(date, dateadd(mm, 1, dateadd(dd, -day(t.RMoth)+1, t.RMoth)), 120), convert(date, t00.BDate, 120)) [Days0]
             from U_RZQR t left join t00 on t00.BaseEntry = t.DocEntry
             where t.DocType = 'Q')
             , t01 as (select a.BCardName, a.CardName, max(a.Days0) maxPeriod, avg(a.Days0) avgPeriod from a group by a.BCardName, a.CardName having max(a.Days0) is not null)
            , t02 as (SELECT T0.CardName 供应商,convert(date,T0.SDate,120) 系统账期开始日期,
            T0.BCardName 酒店,LEFT(Convert(varchar(100),T0.RMoth,102),7) as 货款月份,
            convert(date, dateadd(mm, 1, dateadd(dd, -day(t0.RMoth)+1, t0.RMoth)), 120) 实际账期开始日期,
            convert(date, T0.EDate, 120) 账期结束日期,
            t1.LineTotal 融资金额,t2.FTotal 放款金额
            FROM U_RZQR T0
            INNER JOIN
            (
                SELECT A.DocEntry,SUM(A.LineTotal) as LineTotal,SUM(Convert(numeric(19,2),A.LineTotal*B.RZRate))
              as FKTotal
              FROM U_RZQR1 A
              INNER JOIN U_RZQR B ON A.DocEntry=B.DocEntry
              GROUP BY A.DocEntry
            ) T1 ON T0.DocEntry=T1.DocEntry
            INNER JOIN
            (
                SELECT T0.BaseEntry,SUM(T0.FTotal) as FTotal,MAX(T0.DocDate) as FDate FROM U_OPFK1 T0
                INNER JOIN U_OPFK T1 ON T0.DocEntry=T1.DocEntry
                INNER JOIN U_FKJL T2 ON T2.BaseEntry=T1.DocEntry AND T2.BaseLine=T0.LineNum
                WHERE T1.DocType='F' AND T2.Shenh='Y'
                GROUP BY T0.BaseEntry
            ) T2 ON T0.DocEntry=T2.BaseEntry
            LEFT JOIN
            (
                SELECT T0.BaseEntry,SUM(HGTotal) as HTotal
                FROM U_OPFK1 T0
                INNER JOIN U_OPFK T1 ON T0.DocEntry=T1.DocEntry
                WHERE T1.DocType='G'
                GROUP BY T0.BaseEntry
            ) T3 ON T0.DocEntry=T3.BaseEntry
            LEFT JOIN
            (
                SELECT T0.BaseEntry,SUM(BTotal) as HBTotal,MAX(BDate) as HBDate
                FROM U_OPFK1 T0
                INNER JOIN U_OPFK T1 ON T0.DocEntry=T1.DocEntry
                WHERE T1.DocType='B'
                GROUP BY T0.BaseEntry
            ) T4 ON T0.DocEntry=T4.BaseEntry
            INNER JOIN T_OCRD T5 ON T0.CardCode=T5.CardCode
            WHERE --T1.LineTotal>ISNULL(T4.HBTotal,0)+ISNULL(T3.HTotal,0) ---融资总计大于回款+回购金额
             (ISNULL(T3.BaseEntry,'')='' AND ISNULL(T4.BaseEntry,'')='')
            AND T0.DocType='Q'
            )
            select t02.供应商,t02.酒店, t02.系统账期开始日期, t02.货款月份, t02.实际账期开始日期,
            t02.账期结束日期,
            dateadd(dd, t01.maxPeriod, t02.实际账期开始日期) '关注日期[最大账期]',
            dateadd(dd, t01.avgPeriod, t02.实际账期开始日期) '关注日期[平均账期]',
            DATEDIFF(dd, convert(date,getdate(),120), dateadd(dd, t01.maxPeriod, t02.实际账期开始日期)) '逾期天数[最大账期]',
            DATEDIFF(dd, convert(date,getdate(),120), dateadd(dd, t01.avgPeriod, t02.实际账期开始日期)) '逾期天数[平均账期]',
            convert(numeric(18,2),t02.融资金额) 融资金额, convert(numeric(18,2),t02.放款金额) 放款金额
            from t02 left join t01 on t01.CardName = t02.供应商 and t01.BCardName = t02.酒店 where t01.maxPeriod is not null
            order by 9
            '''
    sqlData, field1 = read_sql(server, userName, passWord, dataBase, sql1)
    write_to_excel(excel, sqlData, field1)
    print 'Done!'
