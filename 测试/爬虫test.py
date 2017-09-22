# -*- coding:utf8 -*-
import urllib2
import re


def readweb(url):
    # values = {'username': 'username', 'password': 'xxx'}
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; rv:2.0.1) Gecko/20100101 Firefox/4.0.1'}
    # data1 = urllib.urlencode(values)
    # getUrl = url
    request = urllib2.Request(url, headers=headers)
    website = urllib2.urlopen(request)
    return website.read()


def getgroup(patten, html):
    items = re.findall(patten, html)
    groupList = []
    for item in items:
        nameStart = item.find('"NAME"') + 8
        nameEnd = item.find('"PARENT_CODE"') - 2
        reCodeStart = item.find('"REGION_CODE":') + 15
        reCodeEnd = item.find('"SHORT_NAME":') - 2
        codeStart = item.find('"CODE":') + 8
        codeEnd = item.find('"}')
        groupName = item[nameStart: nameEnd]
        groupCode = item[codeStart: codeEnd]
        reCode = item[reCodeStart: reCodeEnd]
        groupList.append([groupName, groupCode, reCode])
    return groupList


if __name__ == '__main__':
    try:
        webUrl1 = 'http://wsbs.sz.gov.cn/shenzhen/open'
        # webPatten = re.compile('''{"SHORT_CODE":".{2}","NAME":".*?","PARENT_CODE":".*?","ID":".*?","SORT_ORDER":.*?,"REGION_CODE":".*?","SHORT_NAME":".*?","TYPE":".*?","CODE":".*?"}''', re.S)
        webPatten = re.compile('''{"SHORT_CODE":".{2}",.*?"REGION_CODE":".*?",.*?"CODE":".*?"}''', re.S)
        htmlCode = readweb(webUrl1)
        groupData = getgroup(webPatten, htmlCode)
        webUrl2 = 'http://wsbs.sz.gov.cn/shenzhen/icity/open/type?dept_id=%s&region=%s'
    except urllib2.URLError, e:
        if hasattr(e, "code"):
            print e.code
        if hasattr(e, "reason"):
            print e.reason
