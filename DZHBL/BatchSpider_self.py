# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用: 批量查询企查查企业信息
import time
import requests
import xlrd
import xlsxwriter
from bs4 import BeautifulSoup
import os


# 企查查网站爬虫类
class EnterpriseInfoSpider:
    def __init__(self):
        # 文件相关
        self.excelPath = '../../file/enterprise_data.xlsx'
        self.writePath = '../../file/result_data_%s.xlsx' % time.strftime('%Y%m%d')
        self.sheetName = 'details'
        self.workbook = None
        self.table = None
        self.beginRow = None
        self.worksheet = None
        self.companylist = None
        self.alldata = []
        # 目录页
        self.catalogUrl = "https://www.qichacha.com/search?key="
        # self.catalogUrl = "http://www.qichacha.com/search_index"
        # 详情页（前缀+firmXXXX+后缀）
        self.detailsUrl = "https://www.qichacha.com"

        self.cookie = raw_input("请输入cookie:").decode("gbk").encode("utf-8")
        self.host = "www.qichacha.com"
        self.userAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.106 Safari/537.36"

        # 数据字段名9个
        self.fields = [u'公司名称', u'网站链接', u'注册资本', u'实缴资本', u'经营状态', u'成立日期', u'法律诉讼', u'自身风险', u'关联风险']

    # 爬虫开始前的一些预处理
    def init(self):
        if not os.path.exists('../file'):
            os.mkdir('../file')
        try:
            # 试探是否有该excel文件，#获取行数：workbook.sheets()[0].nrows
            readworkbook = xlrd.open_workbook(self.excelPath)
            self.beginRow = readworkbook.sheets()[0].nrows  # 获取行数
            self.worksheet = readworkbook.sheet_by_index(0)
            self.companylist = self.worksheet.col_values(0, start_rowx=1)

        except Exception, e:
            print e
            self.workbook = xlsxwriter.Workbook(self.excelPath)
            self.table = self.workbook.add_worksheet(self.sheetName)

            # 创建表头字段
            col = 0
            for field in self.fields:
                self.table.write(0, col, field)
                col += 1

            self.workbook.close()
            self.beginRow = 1
            print "已在当前目录下创建enterprise_data.xlsx数据表"

    # 爬虫开始
    def start_spider(self):
        headers = {"Host": 'www.qichacha.com',
                   "User-Agent": r'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:55.0) Gecko/20100101 Firefox/55.0',
                   "Accept": '*/*',
                   "Accept-Language": 'zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3',
                   "Accept-Encoding": 'gzip, deflate',
                   "Referer": 'http://www.qichacha.com/',
                   "Cookie": self.cookie if self.cookie else r'UM_di**********1',
                   "Connection": 'keep-alive',
                   "If-Modified-Since": 'Wed, 30 **********',
                   "If-None-Match": '"59*******"',
                   "Cache-Control": 'max-age=0', }
        # [u'公司名称', u'网站链接', u'注册资本', u'实缴资本', u'经营状态', u'成立日期', u'法律诉讼', u'自身风险', u'关联风险']
        for item in self.companylist:
            # keyword = {"key": item.encode('utf8'), "index": "0", "p": 1}
            url = (self.catalogUrl + item).encode('utf8')
            # driver.get(url)
            response = requests.get(url, headers=headers)
            soup = BeautifulSoup(response.text, 'html.parser')
            href_list = soup.select('.ma_h1')
            if href_list:
                href = href_list[0].attrs['href'].encode('utf8')
            else:
                continue
            detail_url = self.detailsUrl + href
            detail_response = requests.get(detail_url, headers=headers)
            detail_soup = BeautifulSoup(detail_response.text, 'html.parser')
            company = soup.select('.ma_h1')[0].text
            print company
            self_risk = 0
            link_risk = 0
            risk_list = detail_soup.find_all(lambda tag: tag.name == 'span' and tag.get('class') == ["text-danger"])
            if len(risk_list) > 1:
                if risk_list[0].previous.strip() == u'\u81ea\u8eab\u98ce\u9669':
                    self_risk = int(risk_list[0].text)
                elif risk_list[0].previous.strip() == u'\u5173\u8054\u98ce\u9669' or risk_list[1].previous.strip() == u'\u5173\u8054\u98ce\u9669':
                    link_risk = int(risk_list[1].text)
            register_cash = detail_soup.select('.ntable td')[7].text.strip()
            real_cash = detail_soup.select('.ntable td')[9].text.strip()
            status = detail_soup.select('.ntable td')[11].text.strip()
            build_date = detail_soup.select('.ntable td')[13].text.strip()
            fax_list = detail_soup.select('.company-nav-head')
            if fax_list:
                fax_num = int(fax_list[1].text.encode('utf8').split(' ')[1])
            else:
                fax_num = 0
            self.alldata.append([company, detail_url, register_cash, real_cash, status, build_date, fax_num, self_risk, link_risk])
            time.sleep(0.3)
            self.write_to_excel()

    def write_to_excel(self):
        wb = xlsxwriter.Workbook(self.writePath)
        ws = wb.add_worksheet('Sheet1')
        ws.write_row(0, 0, self.fields)
        for i in range(0, len(self.alldata)):
            ws.write_row(i + 1, 0, self.alldata[i])
        wb.close()

# 爬虫入口
########
spider = EnterpriseInfoSpider()
spider.init()
spider.start_spider()
print 'Done!'
