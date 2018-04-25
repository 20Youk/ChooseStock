# -*- coding:utf-8 -*-
# Author: Youk.Lin
# 应用: 轮询admin是否有发送邮件，再调用微信同步发送
import itchat
import time
import xlrd
import pymssql
import os


def read_sql(host, user, password, database, num):
    sql = '''select t.Code, t.Subject, t.[To]
              from T_OERS t
              where right(t.Sender, 22) = '[admin@gcfactoring.cn]' and t.Code >= %d
          '''
    conn = pymssql.connect(host, user, password, database, charset='utf8')
    cursor = conn.cursor()
    cursor.execute(sql, [num])
    data = cursor.fetchall()
    conn.close()
    return data


def read_address_book(filepath):
    wb_1 = xlrd.open_workbook(filepath)
    sheet_1 = wb_1.sheet_by_index(0)
    emails = sheet_1.col_values(1, start_rowx=1)
    wechats = sheet_1.col_values(2, start_rowx=1)
    addressdict = dict(zip(emails, wechats))
    return addressdict


if __name__ == '__main__':
    txtPath = 'C:/MyProgram/file/MaxCode.txt'
    filePath = 'C:/MyProgram/file/AddressBook.xlsx'
    addressDict = read_address_book(filePath)
    if not os.path.exists(txtPath):
        with open(txtPath, 'w') as txt:
            txt.write('1')
        txt.close()
    server = 'localhost'
    userName = 'dbreader'
    passWord = '*****'
    dataBase = 'baoli'
    msg = u'您好，我司已将邮件【{subject}】发送至您的邮箱，请查收并确认，谢谢！'
    itchat.auto_login(hotReload=True, enableCmdQR=True)
    time.sleep(2)
    s = 1
    while s == 1:
        with open(txtPath, 'r') as txt:
            codeNum = int(txt.read())
        txt.close()
        sqlData = read_sql(host=server, user=userName, password=passWord, database=dataBase, num=codeNum)
        maxCode = sqlData[0][0]
        for i in range(0, len(sqlData)):
            maxCode = max(sqlData[i][0], maxCode)
            weChat_user = itchat.search_friends(name=addressDict[sqlData[i][2]])[0]['UserName']
            itchat.send(msg=msg.format(subject=sqlData[i][1]), toUserName=weChat_user)
        with open(txtPath, 'w') as txt:
            txt.write(str(maxCode))
        txt.close()
        # time.sleep(10)
        s += 1
    print 'Done'
