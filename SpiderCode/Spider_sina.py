# -*- coding:utf8 -*-
# Author: Youk.Lin
import urllib2
import re


def ReadUrl(url):
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; rv:2.0.1) Gecko/20100101 Firefox/4.0.1'}
    request = urllib2.Request(url, headers=headers)
    website = urllib2.urlopen(request)
    return website.read()

if __name__ == '__main__':
    webSite = 'http://finance.sina.com.cn/'
    htmlCode = ReadUrl(url=webSite)
    patten = re.compile('''<a target="_blank" href=".*?">.*?</a> ''', re.S)
    group = re.findall(patten, htmlCode)
    print group
