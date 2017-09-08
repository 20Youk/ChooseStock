# -*- coding:utf8 -*-
import urllib2
import urllib

try:
    values = {'username': 'username', 'password': 'xxx'}
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; rv:2.0.1) Gecko/20100101 Firefox/4.0.1'}
    data1 = urllib.urlencode(values)
    url = 'http://chuansong.me/finance'
    getUrl = url + '?' + data1
    request = urllib2.Request(url, headers=headers)
    webSite = urllib2.urlopen(request)
    print webSite.read()
except urllib2.URLError, e:
    if hasattr(e, "code"):
        print e.code
    if hasattr(e, "reason"):
        print e.reason
