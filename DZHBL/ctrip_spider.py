# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:
from bs4 import BeautifulSoup
import re
import xlwt
import xlrd
from xlutils.copy import copy
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
import time  # 时间操作库，强制等待sleep需引入
import socket

# driver.set_page_load_timeout(30)
socket.setdefaulttimeout(30)


class CommentContent:
    # 读取武汉酒店列表
    def __init__(self):
        pass

    def read_hotel(self):
        global total_score, hotel_sale, bar_score
        chrome_option = Options()
        chrome_option.add_argument('--headless')
        chrome_option.add_argument('--disable-gpu')
        driver = webdriver.Chrome(chrome_options=chrome_option)
        index_url = 'http://hotels.ctrip.com/hotel/shenzhen30#ctm_ref=hod_hp_sb_lst'
        driver.get(index_url)  # 打开待爬取酒店列表页面
        time.sleep(5)
        # 模拟浏览器勾选酒店星级为四星级、五星级
        driver.find_element_by_id('star-4').click()
        time.sleep(3)
        driver.find_element_by_id('star-5').click()
        time.sleep(3)
        if len(driver.find_elements_by_id('page_info')) == 0:
            driver.refresh()
            time.sleep(5)
        info_list = []
        hotel_list = []  # 存储所有带爬取酒店的信息
        # 获取需要爬取的页数
        page_list = driver.find_element_by_id('page_info').text.split('\n')
        last_page = int(page_list[page_list.index(u'\u4e0b\u4e00\u9875') - 1]) - 1
        for i in range(last_page):
            # 获取该页的酒店列表
            hotel_loc = driver.find_elements_by_class_name('hotel_item')
            count = 0
            page_count = len(hotel_loc)
            print page_count
            while count < page_count:
                try:
                    # 获取酒店列表的html，供beautiful转换后方便提取
                    hotel = hotel_loc[count].get_attribute("innerHTML")
                    hotel_list.append(BeautifulSoup(hotel, "html.parser"))
                    count += 1
                except Exception, e:
                    # print e
                    print 'get hotel html error'
                    # continue
            # 点击下一页按钮
            driver.find_element_by_class_name('c_down').click()
            time.sleep(10)

        count = 0
        hotel_count = len(hotel_list)
        # 遍历酒店列表里的每一条酒店信息
        while count < hotel_count:
            try:
                # 获取每一条酒店的总体评分、用户推荐%数、查看详情的链接地址
                hotel_name = hotel_list[count].findAll("h2", {"class", "hotel_name"})[0].contents[0]['title']
                total_judgement_score = hotel_list[count].findAll("span", {"class": "hotel_value"})[0].get_text()
                hotel_judgement = hotel_list[count].find("span", {"class": "total_judgement_score"}).get_text().split('%')[0]
                hotel_judgement_person = hotel_list[count].find("span", {"class": "hotel_judgement"}).contents[1].get_text()
                hotel_price = hotel_list[count].find("span", {"class": "J_price_lowList"}).get_text()
                detail_href = hotel_list[count].find("a", {"class": "btn_buy"})['href']
                # 构造酒店详情信息的url
                detail_url = 'http://hotels.ctrip.com/' + detail_href
                try:
                    # 进入酒店详情页面
                    print '1-------------'
                    driver.get(detail_url)
                    print "start new page"
                except TimeoutException:
                    print 'time out after 30 seconds when loading page'
                time.sleep(3)
                # 点击酒店点评
                try:
                    WebDriverWait(driver, 5).until(lambda x: x.find_element_by_id("commentTab")).click()
                    # 程序执行到该处使用driver等待的方式并不能进入超时exception，也不能进入设定好的页面加载超时错误？？？？？
                except socket.error:
                    print 'commtab error'
                    time.sleep(10)
                    # driver.execute("acceptAlert") #此行一直出错，浏览器跳出警告框？？？？无法执行任何有关driver 的操作
                    # continue
                    # driver.quit()
                try:
                    time.sleep(3)
                    bar_score = WebDriverWait(driver, 10).until(
                        lambda x: x.find_element_by_xpath("//div[@class='bar_score']"))
                except Exception, e:
                    print 'bbbbbbbbbbbb'
                    print e
                # bar_score = driver.find_element_by_xpath("//div[@class='bar_score']")
                # 对获取内容进行具体的正则提取
                total_score_ptn = re.compile('(.*?)%')
                try:
                    total_score = total_score_ptn.findall(total_judgement_score)[0]
                    # total_score = total_score_ptn.search(total_judgement_score).group(1)
                    hotel_sale_ptn = re.compile(r'\d+')
                    # hotel_sale = hotel_sale_ptnsearch(hotel_judgement).group(1)
                    hotel_sale = hotel_sale_ptn.findall(hotel_judgement)[0]
                except Exception, e:
                    print 'tote error'
                    print e
                # 获取位置、设施、服务、卫生评分
                bar_scores_ptn = re.compile(r"\d.\d")  # 提取字符串中的数字，格式类似于3.4
                bar_scores = bar_scores_ptn.findall(bar_score.text)
                try:
                    loc_score = bar_scores[0]
                    device_score = bar_scores[1]
                    service_score = bar_scores[2]
                    clean_score = bar_scores[3]
                except Exception, e:
                    print '0------'
                    print e
                    continue
                # 将每个酒店的所有数据以元祖形式存储进hotel_info，存储成元组是为了方便后面写入excel，
                # 后将所有酒店信息追加至info_list
                hotel_info = (hotel_name, total_score, hotel_sale, loc_score, device_score, service_score, clean_score)
                info_list.append(hotel_info)
                count += 1
                # 每一页有25个酒店，每爬取一页显示next page提示，并调用save_score方法存储进excel.
                # 另外重启浏览器，以防止其崩溃
                if count % 24 == 0:
                    print "next page"
                    CommentContent().save_score(info_list)
                    info_list = []
                    driver.close()
                    time.sleep(10)
                    driver = webdriver.Firefox()
            except Exception, e:
                print 'get detail info error'
                # print e
                count += 1
                continue
                # traceback.print_exc()
                # driver.close()
        return info_list

    # 建立数据存储的excel和格式，以及写入第一行
    def build_excel(self):
        wb = xlwt.Workbook()
        sheet1 = wb.add_sheet(u'sheet1', cell_overwrite_ok=True)
        head_style = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
                                 num_format_str='#,##0.00')
        row0 = ('hotel_name', 'total_score', 'sale', 'loc_score', 'device_score', 'service_score', 'clean_score')
        for i in range(0, len(row0)):
            sheet1.write(0, i, row0[i], head_style)
        wb.save('../file/score1.xls')

    # 将数据追加至已建立的excel文件中

    def save_score(self, info_list):
        score_list = info_list
        rb_file = xlrd.open_workbook('../file/score1.xls')
        nrows = rb_file.sheets()[0].nrows  # 获取已存在的excel表格行数
        wb_file = copy(rb_file)  # 复制该excel
        sheet1 = wb_file.get_sheet(0)  # 建立该excel的第一个sheet副本
        try:
            for i in range(0, len(score_list)):
                for j in range(0, len(score_list[i])):
                    # 将数据追加至已有行后
                    sheet1.write(i + nrows, j, score_list[i][j])
        except Exception, e:
            print e
        wb_file.save('../file/score1.xls')
        print 'save success'


if __name__ == '__main__':
    CommentContent().build_excel()
    CommentContent().read_hotel()
