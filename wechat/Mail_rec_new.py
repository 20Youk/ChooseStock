# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:读取腾讯企业邮箱发件箱列表
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xlrd
import itchat
import time


def login_mail(url, username, password):
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    driver = webdriver.Chrome(options=chrome_options)
    driver.get(url=url)
    driver.find_element_by_id('inputuin').clear()
    driver.find_element_by_id('inputuin').send_keys(username)
    driver.find_element_by_id('pp').clear()
    driver.find_element_by_id('pp').send_keys(password)
    driver.find_element_by_id('btlogin').click()
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'mainFrameContainer')))
    return driver

if __name__ == '__main__':
    i = 1
    itchat.auto_login(hotReload=True)
    time.sleep(2)
    wb_1 = xlrd.open_workbook('../../file/AddressBook.xlsx')
    sheet_1 = wb_1.sheet_by_index(0)
    emails = sheet_1.col_values(1, start_rowx=1)
    weChats = sheet_1.col_values(2, start_rowx=1)
    addressDict = dict(zip(emails, weChats))
    website = 'https://exmail.qq.com/login'
    userName = 'junyou.lin@gcfactoring.cn'
    passWord = '****'
    urlDriver = login_mail(url=website, username=userName, password=passWord)
    while True:
        print u'正在进行第%d次扫描......'% i
        urlDriver.find_element_by_id('folder_3_td').click()
        time.sleep(2)
        # WebDriverWait(urlDriver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body')))
        urlDriver.switch_to.frame('mainFrame')
        base_receivers = urlDriver.find_elements_by_xpath('//*[@id="frm"]/div[3]'
                                                          '/table/tbody/tr/td[3]/table/tbody/tr/td[2]')
        base_subjects = urlDriver.find_elements_by_xpath('//*[@id="frm"]/div[3]'
                                                         '/table/tbody/tr/td[3]/table/tbody/tr/td[4]')
        base_times = urlDriver.find_elements_by_xpath('//*[@id="frm"]/div[3]/table/tbody/tr/td[3]/table/tbody/tr/td[5]')
        receivers = [item.get_attribute('title') for item in base_receivers]
        subjects = [item.text for item in base_subjects]
        times = [item.text for item in base_times]
        all_list = zip(receivers, subjects, times)
        msg = u"您好,我司已将邮件【{subject}】发送至您邮箱，请注意查收，谢谢！"
        for item_1 in all_list:
            index_1 = item_1[2].find(u'\u79d2')
            index_2 = item_1[2].find(u'\u5206\u949f\u524d')
            if index_1 != -1:
                user = itchat.search_friends(name=addressDict[item_1[0]])[0]['UserName']
                itchat.send(msg=msg.format(item_1[1]), toUserName=user)
            elif index_2 != -1:
                time_num = int(item_1[2][0: index_2])
                if time_num <= 60:
                    user = itchat.search_friends(name=addressDict[item_1[0]])[0]['UserName']
                    itchat.send(msg=msg.format(item_1[1]), toUserName=user)
        urlDriver.switch_to.default_content()
        time.sleep(20)
        i += 1
