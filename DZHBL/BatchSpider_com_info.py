# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用: 批量查询企查查企业信息
import time
import requests
import xlrd
import xlsxwriter
from bs4 import BeautifulSoup
import os
import logging
import sys


log_path = '../../log/alert/qichahca_%s' % time.strftime('%Y%m%d') + '.txt'
# 设置日志输出
reload(sys)
sys.setdefaultencoding('gbk')
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
# 设置日志输出格式
formatter = logging.Formatter('%(asctime)s [%(levelname)s]  %(name)s : %(message)s')
# 设置日志文件路径、告警级别过滤、输出格式
fh = logging.FileHandler(log_path)
fh.setLevel(logging.WARN)
fh.setFormatter(formatter)
# 设置控制台告警级别、输出格式
ch = logging.StreamHandler()
# ch.setLevel(logging.INFO)
ch.setFormatter(formatter)
# 载入配置
logger.addHandler(fh)
logger.addHandler(ch)


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
        self.fields = [u'查询公司', u'公司名称', u'网站链接', u'电话', u'注册资本', u'实缴资本', u'经营状态', u'成立日期', u'公司地址',
                       u'所属区域', u'登记机关', u'公司简介', u'参保人数', u'被执行人信息', u'失信信息', u'商标信息', u'专利信息',
                       u'证书信息', u'作品著作信息', u'软件信息', u'网站信息']

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
        # [u'公司名称', u'网站链接', u'电话', u'注册资本', u'实缴资本', u'经营状态', u'成立日期', u'公司地址',
        #                u'所属区域', u'登记机关', u'公司简介', u'参保人数', u'被执行人信息', u'失信信息', u'商标信息', u'专利信息',
        #                u'证书信息', u'作品著作信息', u'软件信息,' u'网站信息']
        k = 1
        for item in self.companylist:
            try:
                # keyword = {"key": item.encode('utf8'), "index": "0", "p": 1}
                url = (self.catalogUrl + item).encode('utf8')
                # driver.get(url)
                response = requests.get(url, headers=headers)
                soup = BeautifulSoup(response.text, 'html.parser')
                href_list = soup.select('.ma_h1')
                if len(href_list) > 0:
                    href = href_list[0].attrs['href'].encode('utf8')
                else:
                    continue
                detail_url = self.detailsUrl + href
                detail_response = requests.get(detail_url, headers=headers)
                detail_soup = BeautifulSoup(detail_response.text, 'html.parser')
                company = soup.select('.ma_h1')[0].text
                print u'查询第%d个企业信息：%s' % (k, company)
                k += 1
                length = len(detail_soup.select('.company-nav-items'))
                if length < 6:
                    continue
                check = detail_soup.select('.company-nav-items')[0].text.split()
                if check[0] == u'\u80a1\u7968\u884c\u60c5':
                    falv = detail_soup.select('.company-nav-items')[2].text.split()
                    chanquan = detail_soup.select('.company-nav-items')[6].text.split()
                    cominfo = detail_soup.find(id='base_div').find(id='Cominfo').select('.ntable td')
                else:
                    falv = detail_soup.select('.company-nav-items')[1].text.split()
                    chanquan = detail_soup.select('.company-nav-items')[5].text.split()
                    cominfo = detail_soup.select('.ntable td')
                phone = detail_soup.select('.cvlu')[0].text.strip().split(' ')[0]
                zhixing = int(falv[1])
                shixin = int(falv[3])
                shangbiao = int(999 if chanquan[1] == '999+' else chanquan[1])
                zhuanli = int(999 if chanquan[3] == '999+' else chanquan[3])
                zhengshu = int(999 if chanquan[5] == '999+' else chanquan[5])
                zuopin = int(999 if chanquan[7] == '999+' else chanquan[7])
                ruanjian = int(999 if chanquan[9] == '999+' else chanquan[9])
                wangzhan = int(999 if chanquan[11] == '999+' else chanquan[11])
                # self_risk = 0
                # link_risk = 0
                # risk_list = detail_soup.find_all(lambda tag: tag.name == 'span' and tag.get('class') == ["text-danger"])
                # if len(risk_list) > 1:
                #     if risk_list[0].previous.strip() == u'\u81ea\u8eab\u98ce\u9669':
                #         self_risk = int(risk_list[0].text)
                #     elif risk_list[0].previous.strip() == u'\u5173\u8054\u98ce\u9669' or risk_list[1].previous.strip() == u'\u5173\u8054\u98ce\u9669':
                #         link_risk = int(risk_list[1].text)
                register_cash = cominfo[7].text.strip()
                real_cash = cominfo[9].text.strip()
                status = cominfo[11].text.strip()
                build_date = cominfo[13].text.strip()
                address = cominfo[43].text.strip().split("\n")[0]
                area = cominfo[31].text.strip()
                register_gov = cominfo[29].text.strip()
                description = cominfo[45].text.strip()
                canbao0 = cominfo[37].text.strip()
                canbao = 0 if canbao0 == '-' else int(canbao0)
                # fax_list = detail_soup.select('.company-nav-head')
                # if fax_list:
                #     fax_num = int(fax_list[1].text.encode('utf8').split(' ')[1])
                # else:
                #     fax_num = 0
                # self.alldata.append([company, detail_url, register_cash, real_cash, status, build_date, fax_num, self_risk, link_risk])
            # [u'公司名称', u'网站链接', u'电话', u'注册资本', u'实缴资本', u'经营状态', u'成立日期', u'公司地址',
            #                u'所属区域', u'公司简介', u'参保人数', u'被执行人信息', u'失信信息', u'商标信息', u'专利信息',
            #                u'证书信息', u'作品著作信息', u'软件信息,' u'网站信息']
                self.alldata.append([item, company, detail_url, phone, register_cash, real_cash, status, build_date, address, area, register_gov, description, canbao,
                                     zhixing, shixin, shangbiao, zhuanli, zhengshu, zuopin, ruanjian, wangzhan])
                time.sleep(1)
                self.write_to_excel()
            except (Exception, IOError), e:
                logger.error(e, exc_info=True)
                pass
            continue

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
