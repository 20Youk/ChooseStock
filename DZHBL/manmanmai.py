# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:
import urllib2
import requests
import json

url = 'http://sapi.manmanbuy.com/Search.aspx?AppKey=nAtYiJJ80XtmHs28&Key=%s&Class=0&Brand=0&Site=0&PriceMin=0&PriceMax=0&PageNum=1&PageSize=30&OrderBy=price&ZiYing=false&ExtraParameter=0'
key = u'益而高黑色长尾票夹no.TY145'
url_key = url % key
request = urllib2.Request(url_key.encode('GB2312'))
response = urllib2.urlopen(request)
content = response.read()
# content = requests.get(url_key.encode('GB2312')).content
retDict = json.loads(content.decode('gbk'))
print retDict
