# -*- coding:utf8 -*-
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

allData = []
url = 'http://www.haodf.com/doctor/DE4r0Fy0C9LuGYKIlfMCIlReozGolC54J.htm'
chrome_option = Options()
chrome_option.add_argument('--headless')
chrome_option.add_argument('--disable-gpu')
driver = webdriver.Chrome(chrome_options=chrome_option)
driver.get(url)
about_1 = driver.find_elements_by_xpath('//*[@id="bp_doctor_about"]/div/div[2]/div/table[1]/tbody/tr')
group_1 = about_1[-4].text.split('\n')[1]
title_list = about_1[-3].text.split(u'\uff1a')
if len(title_list) > 1:
    title_1 = about_1[-3].text.split(u'\uff1a')[1]
else:
    title_1 = u'无'
des_1 = about_1[-2].text.split(u'\uff1a')[1]
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

print [group_1.encode('utf8'), title_1.encode('utf8'), score_1.encode('utf8'), num_1.encode('utf8'), num_2.encode('utf8')]

