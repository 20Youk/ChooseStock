# -*- coding:utf8 -*-
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait


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
        # 代码来源：页面审查元素，通过元素路径定位
        base_msg = self.driver.find_elements_by_xpath('//span[@class="baseinfo"]/a')
        wait = WebDriverWait(self.driver, 2)
        wait.until(lambda driver: driver.find_element_by_xpath('//div[@id="footer"]'))
        for item in base_msg:
            print u'服务事项名称: %s, 网址: %s\n' %(item.text, item.get_attribute('href').encode('utf8'))
            brief.append([item.text, item.get_attribute('href').encode('utf8')])
        pages = self.driver.find_elements_by_xpath('//div[@class="pages"]/a')
        nextpagetext = pages[-1].text
        if pages:
            for i in range(1, len(pages) - 1):
                self.driver.find_element_by_link_text(nextpagetext).click()
                wait = WebDriverWait(self.driver, 2)
                wait.until(lambda driver: driver.find_element_by_xpath('//div[@class="pages"]/a'))
                base_msg = self.driver.find_elements_by_xpath('//span[@class="baseinfo"]/a')
                for item in base_msg:
                    print u'服务事项名称: %s, 网址: %s\n' %(item.text, item.get_attribute('href').encode('utf8'))
                    brief.append([item.text, item.get_attribute('href').encode('utf8')])
        return brief

if __name__ == '__main__':
    spider = Spider()
    print spider.getdetailpage('http://wsbs.sz.gov.cn/shenzhen/icity/open/type?dept_id=693956469&region=440300')
