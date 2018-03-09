# -*- coding:utf8 -*-
# Author: Youk.Lin
import urllib2
import re
from selenium_test import webdriver
from selenium_test.webdriver.support.ui import WebDriverWait
import xlsxwriter
import datetime
import sys


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


class Spider:
    def __init__(self):
        self.page = 1
        self.dirName = 'MMSpider'
        # 这是一些配置 关闭loadimages可以加快速度 但是第二页的图片就不能获取了打开(默认)
        cap = webdriver.DesiredCapabilities.PHANTOMJS
        cap["phantomjs.page.settings.resourceTimeout"] = 1000
        # cap["phantomjs.page.settings.loadImages"] = False
        # cap["phantomjs.page.settings.localToRemoteUrlAccessEnabled"] = True
        self.driver = webdriver.PhantomJS(desired_capabilities=cap)

    def getdetailpage(self, url):
        self.driver.get(url)
        # base_msg = self.driver.find_elements_by_xpath('//div[@class="mm-p-info mm-p-base-info"]/ul/li')
        brief = []
        base_msg = self.driver.find_elements_by_xpath('//span[@class="baseinfo"]/a')
        wait = WebDriverWait(self.driver, 10)
        wait.until(lambda driver: driver.find_element_by_xpath('//div[@id="footer"]'))
        for item in base_msg:
            print u'服务事项名称: {0:s}, 网址: {1:s}\n'.format(item.text, item.get_attribute('href').encode('utf8'))
            brief.append([item.text, item.get_attribute('href').encode('utf8')])
        pages = self.driver.find_elements_by_xpath('//div[@class="pages"]/a')
        if pages:
            nextpagetext = pages[-1].text
            for i in range(1, len(pages) - 1):
                self.driver.find_element_by_link_text(nextpagetext).click()
                wait = WebDriverWait(self.driver, 2)
                wait.until(lambda driver: driver.find_element_by_xpath('//div[@class="pages"]/a'))
                base_msg = self.driver.find_elements_by_xpath('//span[@class="baseinfo"]/a')
                for item in base_msg:
                    print u'服务事项名称: {0:s}, 网址: {1:s}\n'.format(item.text, item.get_attribute('href').encode('utf8'))
                    brief.append([item.text, item.get_attribute('href').encode('utf8')])
        return brief


def writetoexcel(filepath, bookname, sheetname, field, data):
    today = datetime.datetime.now().strftime('%Y%m%d')
    wb = xlsxwriter.Workbook(filepath + bookname + today + '.xlsx')
    sheet = wb.add_worksheet(sheetname)
    sheet.write_row(0, 0, field)
    for i in range(0, len(field)):
        for row in range(0, len(data)):
            sheet.write(row + 1, i, data[row][i])
    wb.close()


def alldata(html):
    # 目录名称
    webPatten0 = re.compile('''<div title="" class="guide-title">.*?</div>''', re.S)
    result0 = re.findall(webPatten0, html)
    if result0:
        item0 = result0[0].split('\r\n')[2].strip()
        # 事项类型
        webPatten1 = re.compile('''<td colspan="2" class="td_1">.*?</td>''', re.S)
        result1 = re.findall(webPatten1, html)
        item1 = result1[0].split('\r\n')[2].strip()
        # 部门名称
        webPatten2 = re.compile('''<td colspan="2" class="td_2">.*?</td>''', re.S)
        result2 = re.findall(webPatten2, html)
        item2 = result2[0].split('>')[1][:-4]
        # 行使层级
        webPatten3 = re.compile('''<td colspan="2" class="td_3 formatSscj" >.*?</td>''', re.S)
        result3 = re.findall(webPatten3, html)
        item3 = result3[0].split('\t')[2]
        # 基本编码
        webPatten4 = re.compile('\xe5\x9f\xba\xe6\x9c\xac\xe7\xbc\x96\xe7\xa0\x81</th>.*?<td colspan="2">.*?</td>', re.S)
        result4 = re.findall(webPatten4, html)
        item4 = result4[0].split('>')[2][:-4]
        # 办理形式
        webPatten5 = re.compile('\xe5\x8a\x9e\xe7\x90\x86\xe5\xbd\xa2\xe5\xbc\x8f</th>.*?<td colspan="2">.*?</td>', re.S)
        result5 = re.findall(webPatten5, html)
        item5 = result5[0].split('\t')[1][:-2]
        # 实施主体性质
        webPatten6 = re.compile('\xe5\xae\x9e\xe6\x96\xbd\xe4\xb8\xbb\xe4\xbd\x93\xe6\x80\xa7\xe8\xb4\xa8</th>.*?<td colspan="2">.*?</td>', re.S)
        result6 = re.findall(webPatten6, html)
        item6 = result6[0].split('\r\n')[4].strip()
        # 材料信息
        webPatten7 = re.compile('<td>.{1,2}</td>.*?<td class="td-info">.*?</tr>', re.S)
        result7 = re.findall(webPatten7, html)
        # resultList = [部门名称, 目录名称, 事项类型, 基本编码, 办理形式, 行使层级, 实施主体性质, 材料名称, 原件份数, 复印件份数, 纸质/电子版]
        resultList = []
        for j in range(0, len(result7)):
            items = result7[j].split('\r\n')
            # 材料名称
            item71 = items[1].strip()[4: -5]
            # 原件份数
            item72 = int(items[8].strip()[4: -5])
            # 复印件份数
            item73 = int(items[9].strip()[4: -5])
            # 纸质/电子版
            item74 = items[12].strip()
            resultList.append([item2, item0, item1, item4, item5, item3, item6, item71, item72, item73, item74])
        return resultList
    else:
        return []


if __name__ == '__main__':
    try:
        reload(sys)
        sys.setdefaultencoding('utf-8')
        webUrl1 = 'http://wsbs.sz.gov.cn/shenzhen/open'
        # webPatten = re.compile('''{"SHORT_CODE":".{2}","NAME":".*?","PARENT_CODE":".*?","ID":".*?","SORT_ORDER":.*?,"REGION_CODE":".*?","SHORT_NAME":".*?","TYPE":".*?","CODE":".*?"}''', re.S)
        webPatten = re.compile('''{"SHORT_CODE":".{2}",.*?"REGION_CODE":".*?",.*?"CODE":".*?"}''', re.S)
        htmlCode = readweb(webUrl1)
        # groupData = [[groupName, groupCode, reCode]]
        groupData = getgroup(webPatten, htmlCode)
        webUrl2 = 'http://wsbs.sz.gov.cn/shenzhen/icity/open/type?dept_id=%s&region=%s'
        urlList = []
        spider = Spider()
        for jtem in groupData:
            urlList.extend(spider.getdetailpage(webUrl2 % (jtem[1], jtem[2])))
        filePath = '../../excel/'
        bookName = 'ReportURL'
        sheetName = u'服务事项名称及网址'
        fields = [u'服务事项名称', u'网址']
        writetoexcel(filepath=filePath, bookname=bookName, sheetname=sheetName, field=fields, data=urlList)
        todayStr = datetime.datetime.now().strftime('%Y%m%d')
        workBook = xlsxwriter.Workbook('../../excel/AllData%s.xlsx' % todayStr)
        workSheet = workBook.add_worksheet(u'公开信息')
        fields = [u'部门名称', u'目录名称', u'事项类型', u'基本编码', u'办理形式', u'行使层级', u'实施主体性质', u'材料名称', u'原件份数', u'复印件份数', u'纸质/电子版']
        workSheet.write_row(0, 0, fields)
        logFile = open('../../log/log_%s.log' % todayStr, mode='a', buffering=1)
        k = 0
        for u in range(0, len(urlList)):
            print u'第{0:d}次爬虫获取机构【{1:s}】网址【{2:s}】的相关信息...'.format(u + 1, urlList[u][0], urlList[u][1])
            logFile.write(u'\n{0:s} : 正在通过爬虫获取机构【{1:s}】网址【{2:s}】的相关信息...'.format(
                datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                urlList[u][0], urlList[u][1]))
            htmlCode = readweb(urlList[u][1])
            oneData = alldata(htmlCode)
            for n in range(0, len(oneData)):
                workSheet.write_row(n + k + 1, 0, oneData[n])
            k += len(oneData)
            logFile.write(u'\n{0:s} : 成功通过爬虫获取机构【{1:s}】网址【{2:s}】的相关信息...'.format(
                datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                urlList[u][0], urlList[u][1]))
        workBook.close()
        logFile.write(u'\n{0: s}: 执行完成!!!'.format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
        logFile.close()
    except (urllib2.URLError, Exception, IOError), e:
        print u'程序运行出错,请查看日志'
        todayStr = datetime.datetime.now().strftime('%Y%m%d')
        logFile = open('../../log/log_%s.log' % todayStr, mode='a', buffering=1)
        logFile.write(u'\n{0:s} : {1:s}'.format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), e))
        logFile.close()
