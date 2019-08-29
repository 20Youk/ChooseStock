# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:
import urllib2
import requests
import json

url = 'http://api.taokezhushou.com/api/v1/search?app_key=3e138417ccf70a0c&q=%s'
key = u'黑色长尾票夹'
url_key = url % key
request = urllib2.Request(url_key.encode('utf8'))
response = urllib2.urlopen(request)
content = response.read()
# content = requests.get(url_key.encode('GB2312')).content
retDict = json.loads(content)
print retDict
