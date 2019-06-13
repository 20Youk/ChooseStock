# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:
from bs4 import BeautifulSoup
import xlwt
import xlrd
from xlutils.copy import copy
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time  # 时间操作库，强制等待sleep需引入
import socket

# driver.set_page_load_timeout(30)
socket.setdefaulttimeout(30)

star_list = [ 'star5']
star_dict = {'star3': u'三星级',
             'star4': u'四星级',
             'star5': u'五星级'}
# area_list = ['hangzhou17', 'chongqing4', 'suzhou14', "xi'an10", 'sanya43', 'wuhan477', 'nanjing12', 'guilin33',
#              'changsha206', 'ningbo375', 'qingdao7', 'dongguan223', 'xiamen25', 'tianjin3', 'shenzhen30',
#              'guangzhou32', 'foshan251', 'huizhou299', 'zhuhai31']
area_list = ['zhongshan553', 'jiangmen316', 'zhaoqing552']
file_path = '../file/hotel_list.xls'


class CommentContent:
    # 读取武汉酒店列表
    def __init__(self):
        pass

    def read_hotel(self):
        global info_list
        chrome_option = Options()
        chrome_option.add_argument('--headless')
        chrome_option.add_argument('--disable-gpu')
        index_url = 'http://hotels.ctrip.com/hotel/%s/%s'
        count = 1
        for area in area_list:
            for star in star_list:
                try:
                    driver = webdriver.Chrome(chrome_options=chrome_option)
                    driver.implicitly_wait(10)
                    driver.get(index_url % (area, star))  # 打开待爬取酒店列表页面
                    # driver.maximize_window()
                    title = driver.find_element_by_class_name('path_bar').find_element_by_tag_name('h1').text
                    print u'获取%s的信息........' % title
                    info_list = []
                    # 获取需要爬取的页数
                    page_list = driver.find_element_by_id('page_info').text.split('\n')
                    last_page = int(page_list[page_list.index(u'\u4e0b\u4e00\u9875') - 1])
                    hotel_area = driver.find_element_by_id('txtCity').get_attribute('value')
                    hotel_star = star_dict[star]  # 星级
                    for i in range(0, last_page):
                        # 获取该页的酒店列表
                        hotel_loc = driver.find_elements_by_class_name('hotel_item')
                        if len(hotel_loc) == 0:
                            continue
                        for hotel in hotel_loc:
                            try:
                                # 获取酒店列表的html，供beautiful转换后方便提取
                                hotel = hotel.get_attribute("innerHTML")
                                hotel_detail = BeautifulSoup(hotel, "html.parser")
                                hotel_name = hotel_detail.findAll("h2", {"class", "hotel_name"})[0].contents[0]['title']
                                print u'获取第%d家酒店#%s#的基本信息.....' % (count, hotel_name)
                                total_judgement_score_0 = hotel_detail.findAll("span", {"class": "hotel_value"})
                                if len(total_judgement_score_0) == 0:
                                    total_judgement_score = 0
                                else:
                                    total_judgement_score = total_judgement_score_0[0].get_text()  # 总评分
                                hotel_judgement = hotel_detail.find("span", {"class": "total_judgement_score"}).get_text().split('%')[0]  # 用户推荐比例
                                hotel_judgement_person = hotel_detail.find("span", {"class": "hotel_judgement"}).contents[1].get_text()  # 住客点评数
                                hotel_price = hotel_detail.find("span", {"class": "J_price_lowList"}).get_text()  # 最低房价
                                detail_href = hotel_detail.find("a", {"class": "btn_buy"})['href']
                                # 构造酒店详情信息的url
                                detail_url = 'http://hotels.ctrip.com/' + detail_href   # 酒店链接
                                # 将每个酒店的所有数据以元组形式存储进hotel_info，存储成元组是为了方便后面写入excel，
                                # 后将所有酒店信息追加至info_list
                                hotel_info = (hotel_name, detail_url, hotel_area, hotel_star, float(total_judgement_score), int(hotel_judgement), int(hotel_judgement_person), int(hotel_price))
                                info_list.append(hotel_info)
                                count += 1
                            except Exception, e:
                                print e
                                print 'get hotel html error'
                                # continue
                        # 进入下一页
                        if i < last_page - 1:
                            btn_next = driver.find_element_by_id('txtpage')
                            btn_next.clear()
                            btn_next.send_keys(i + 2)
                            driver.find_element_by_class_name('c_page_submit').click()
                            time.sleep(10)
                    # 每爬取一次地区+星级酒店显示next work提示，并调用save_score方法存储进excel.
                    # 另外重启浏览器，以防止其崩溃
                    print "next work"
                    CommentContent().save_score(info_list)
                    driver.close()
                    time.sleep(2)
                except (Exception, IOError), e:
                    pass
                    print e
                finally:
                    print time.ctime()
        return info_list

    # 建立数据存储的excel和格式，以及写入第一行
    def build_excel(self):
        wb = xlwt.Workbook()
        sheet1 = wb.add_sheet(u'sheet1', cell_overwrite_ok=True)
        head_style = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
                                 num_format_str='#,##0.00')
        row0 = (u'酒店名称', u'链接', u'地区', u'酒店星级', u'总分', u'推荐用户比例[单位:%]', u'点评用户数', u'最低房价')
        for i in range(0, len(row0)):
            sheet1.write(0, i, row0[i], head_style)
        wb.save(file_path)

    # 将数据追加至已建立的excel文件中

    def save_score(self, score_list):
        rb_file = xlrd.open_workbook(file_path)
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
        wb_file.save(file_path)
        print 'save success'


if __name__ == '__main__':
    CommentContent().build_excel()
    CommentContent().read_hotel()
