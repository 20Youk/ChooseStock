# -*-coding:utf8-*-
import itchat
import time

itchat.auto_login(hotReload=True)
time.sleep(1)
user = itchat.search_friends(name=u'陆泓利')[0]['UserName']
itchat.send(msg=u"test", toUserName=user)
itchat.logout()
