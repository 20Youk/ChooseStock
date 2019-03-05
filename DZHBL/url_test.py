# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time  # 时间操作库，强制等待sleep需引入


def read_hotel(index_url):
    chrome_option = Options()
    # chrome_option.add_argument('--headless')
    chrome_option.add_argument('--disable-gpu')
    driver = webdriver.Chrome(chrome_options=chrome_option)
    driver.get(index_url)  # 打开待爬取酒店列表页面
    driver.maximize_window()
    hotel_name = driver.find_element_by_id('txtCity')
    hotel_name.clear()
    hotel_name.send_keys(u'\u6df1\u5733')
    search_button = driver.find_element_by_id('btnSearch')
    search_button.click()
    time.sleep(5)
    return driver

web = 'http://hotels.ctrip.com/hotel/beijing1#ctm_ref=hod_hp_sb_lst'
web_detail = read_hotel(web)
print web_detail
