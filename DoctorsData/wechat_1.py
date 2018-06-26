# -*- coding:utf-8 -*-
# Author: Youk.Lin
# 应用: 调用微信将对应的工资表发送给员工
import itchat
import time
import xlrd
import pandas


def read_address_book(filepath):
    wb_1 = xlrd.open_workbook(filepath)
    sheet_1 = wb_1.sheet_by_index(0)
    emails = sheet_1.col_values(1, start_rowx=1)
    wechats = sheet_1.col_values(2, start_rowx=1)
    addressdict = dict(zip(emails, wechats))
    return addressdict


if __name__ == '__main__':
    filePath = 'C:/MyProgram/file/testbook.xlsx'
    addressDict = pandas.read_excel(filePath, sheet_name=0)
    field = list(addressDict.columns)
    weChat_userName = field[1]
    field.pop(1)
    # addressDict = read_address_book(filePath)
    msg = u'姓名：%s \n月份：%s\n金额1：%.2f\n金额2：%.2f\n金额3：%.2f\n金额4：%.2f'
    itchat.auto_login(hotReload=True, enableCmdQR=False)
    time.sleep(2)
    for i in range(0, len(addressDict)):
        data = []
        for item in field:
            data.append(addressDict[item][i])
        msg_1 = msg % tuple(data)
        weChat_user = itchat.search_friends(name=addressDict[weChat_userName][i])[0]['UserName']
        itchat.send(msg=msg_1, toUserName=weChat_user)
    print 'Done'
