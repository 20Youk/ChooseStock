# -*- coding:utf8 -*-
# Author: Youk.Lin
import urllib2
import re
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import xlsxwriter
import datetime
import sys

# 获取医生和网址
site_list = []
try:
    reload(sys)
    sys.setdefaultencoding('utf-8')
    for jj in range(1, 9):
        url = 'http://haoping.haodf.com/keshi/2009000/daifu_guangdong_' + str(jj) + '.htm'
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; rv:2.0.1) Gecko/20100101 Firefox/4.0.1'}
        request = urllib2.Request(url, headers=headers)
        website = urllib2.urlopen(request)
        website.encoding = 'gb2312'
        sites = re.findall('<a href="(.*?)" target="_blank" class="blue">.*?</a>', website.read())
        # with open('C:\Users\Administrator\Desktop\HDF\doctors.txt', 'a') as textFile:
        #     for ii in range(0, len(sites)):
        #         textFile.write(sites[ii][0] + ',' + sites[ii][1] + '\n')
        for ii in range(0, len(sites)):
            site_list.append(sites[ii])
    # 获取医生个人信息
    chrome_option = Options()
    chrome_option.add_argument('--headless')
    chrome_option.add_argument('--disable-gpu')
    driver = webdriver.Chrome(chrome_options=chrome_option)
    allData = []
    for item in site_list:
        url = item
        driver.get(url)
        name_1 = driver.find_element_by_xpath('//*[@id="doctor_header"]/div[1]/div/a/h1/span[1]').text
        about_1 = driver.find_elements_by_xpath('//*[@id="bp_doctor_about"]/div/div[2]/div/table[1]/tbody/tr')
        group_1 = about_1[-4].text.split('\n')[1]
        title_list = about_1[-3].text.split(u'\uff1a')
        if len(title_list) > 1:
            title_1 = about_1[-3].text.split(u'\uff1a')[1]
        else:
            title_1 = u'无'
        des_1 = about_1[-2].text[5:]
        # 推荐热度
        score_1 = driver.find_element_by_class_name('r-p-l-score').text
        # 临床经验和患者投票
        num_1ist = driver.find_elements_by_xpath('//*[@id="doctorgood"]/div[1]/table/tbody/tr/td')
        if len(num_1ist) == 1:
            num_1 = num_1ist[0].text
            num_2 = num_1ist[0].text
        elif len(num_1ist) == 2:
            num_1 = num_1ist[0].text
            num_2 = num_1ist[1].text
        else:
            num_1 = u'无'
            num_2 = u'无'
        allData.append([name_1, item, group_1, title_1, des_1, score_1, num_1, num_2])
    # 写入excel
    today = datetime.datetime.now().strftime('%Y%m%d')
    wb_1 = xlsxwriter.Workbook(r'C:\Users\Administrator\Desktop\HDF\data%s.xlsx' % today)
    sheet_1 = wb_1.add_worksheet('Sheet1')
    field = [u'医生', u'主页', u'科室', u'职称', u'擅长', u'推荐热度', u'临床经验', u'患者投票']
    sheet_1.write_row(0, 0, field)
    for row in range(0, len(allData)):
        for col in range(0, len(field)):
            sheet_1.write(row + 1, col, allData[row][col])
    wb_1.close()
    print "Done!"
except (urllib2.URLError, Exception, IOError), e:
    print u'程序运行出错,请查看日志'
    todayStr = datetime.datetime.now().strftime('%Y%m%d')
    logFile = open(r'C:\Users\Administrator\Desktop\HDF\log_%s.log' % todayStr, mode='a', buffering=1)
    logFile.write(u'\n{0:s} : {1:s}'.format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), e))
    logFile.close()
