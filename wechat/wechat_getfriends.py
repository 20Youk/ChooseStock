# -*-coding:utf8-*-
import itchat
import time
import xlsxwriter

itchat.auto_login(hotReload=True)
time.sleep(1)
chatRooms = itchat.get_chatrooms()
print chatRooms
# allFriends = itchat.get_friends()
# print allFriends
# wb = xlsxwriter.Workbook(r'C:\MyProgram\file\WeChat.xlsx')
# sheet = wb.add_worksheet('Sheet1')
# fields = ['UserName', 'NickName', 'RemarkName']
# sheet.write_row(0, 0, fields)
# for i in range(0, len(allFriends)):
#     for j in range(0, len(fields)):
#         sheet.write(i + 1, j, dict(allFriends[i])[fields[j]])
# wb.close()

itchat.logout()
