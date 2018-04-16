# -*-coding:utf8-*-
import itchat
import time
import xlrd
import pymssql


def read_sql(host, user, password, database, num):
    sql = '''select t.Code, t.Subject, t.[To]
              from T_OERS t
              where right(t.Sender, 22) = '[admin@gcfactoring.cn]' and t.Code >= %d
          '''
    conn = pymssql.connect(host, user, password, database, charset='utf8')
    cursor = conn.cursor()
    cursor.execute(sql, [num])
    data = cursor.fetchall()
    return data


def read_address_book(filepath):
    wb_1 = xlrd.open_workbook(filepath)
    sheet_1 = wb_1.sheet_by_index(0)
    emails = sheet_1.col_values(1, start_rowx=1)
    wechats = sheet_1.col_values(2, start_rowx=1)
    addressdict = dict(zip(emails, wechats))
    return addressdict


if __name__ == '__main__':
    while True:
        itchat.auto_login(hotReload=True)
        time.sleep(2)
        user = itchat.search_friends(name=names[0])[0]['UserName']
        itchat.send(msg=u"test", toUserName=user)

