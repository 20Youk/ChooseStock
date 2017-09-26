# -*- coding: utf-8 -*-
# Author: Youk.Lin
import xlrd
import xlsxwriter
from WindPy import *


def tradeanalysis():
    readbook = xlrd.open_workbook(r'../../excel/EntrustData.xls')
    readsheet = readbook.sheet_by_index(0)
    # 获取date,code,stockname,direction,price,time1
    datelist = readsheet.col_values(1, start_rowx=1)
    codelist = readsheet.col_values(8, start_rowx=1)
    namelist = readsheet.col_values(9, start_rowx=1)
    directionlist = readsheet.col_values(10, start_rowx=1)
    pricelist = readsheet.col_values(12, start_rowx=1)
    timelist = readsheet.col_values(37, start_rowx=1)
    lawlist = []
    highlist = []
    lowlist = []
    highprolist = []
    lowprolist = []
    w.start()
    for i in range(0, len(datelist)):
        code = codelist[i].encode('utf8')
        if code[0] == '6':
            windcode = code + '.SH'
        else:
            windcode = code + '.SZ'
        date1 = datelist[i].encode('utf8')
        time1 = timelist[i].encode('utf8')
        if datetime.strptime(time1, '%H:%M:%S') <= datetime.strptime('09:31:00', '%H:%M:%S'):
            datetime1 = date1 + ' 09:31:00'
        else:
            datetime1 = date1 + ' ' + time1
        wsidata = w.wsi(windcode, "high,low", "%s 09:30:00" % date1, datetime1, "")
        if wsidata.ErrorCode == 0:
            maxhigh = max(wsidata.Data[0])
            minlow = min(wsidata.Data[1])
            highproportion = round(round(pricelist[i], 2) / maxhigh - 1, 4)
            abshighproportion = abs(highproportion)
            lowproportion = round(round(pricelist[i], 2) / minlow - 1, 4)
            highlist.append(maxhigh)
            lowlist.append(minlow)
            highprolist.append(highproportion)
            lowprolist.append(lowproportion)
            # 对应Unicode : u'上涨' u'\u4e0a\u6da8' , u'下跌'  u'\u4e0b\u8dcc'
            if lowproportion <= 0.01:
                if abshighproportion <= 0.01:
                    lawlist.append(u'无规律')
                else:
                    lawlist.append(u'\u4e0b\u8dcc' + directionlist[i])
            elif lowproportion > 0.01 >= abshighproportion:
                lawlist.append(u'\u4e0a\u6da8' + directionlist[i])
            elif 0.01 < abshighproportion <= 0.02 <= lowproportion:
                lawlist.append(u'\u4e0a\u6da8' + directionlist[i])
            else:
                lawlist.append(u'无规律')
        else:
            lawlist.append(0)
            highlist.append(0)
            lowlist.append(0)
            highprolist.append(0)
            lowprolist.append(0)
    w.stop()
    today = datetime.now().strftime('%Y%m%d')
    wb = xlsxwriter.Workbook('../../excel/TradeData_%s.xlsx' % today)
    sheet = wb.add_worksheet('Sheet1')
    sheet.set_column('A:A', 11)
    # 标题格式
    wrapstyle = wb.add_format()
    # wrapstyle.set_text_wrap()  # 自动换行
    wrapstyle.set_pattern(1)    # 填充整个单元格
    wrapstyle.set_fg_color('#5C5C5C')   # 单元格底色（RGB颜色代码表）
    wrapstyle.set_font_name(u'宋体')  # 字体
    wrapstyle.set_border(2)     # 设置边框（线条类型）
    wrapstyle.set_font_size(9)      # 字体大小
    wrapstyle.set_align('left')
    wrapstyle.set_align('vcenter')
    # 数字格式
    decimalstyle = wb.add_format({
        'bold': False,   # 字体加粗
        'align': 'left',
        'valign': 'vcenter',
        'num_format': '0.00%',
        'border': 2,
        'font_name': u'宋体',
        'font_size': 9
    })
    # 其余格式
    otherstyle = wb.add_format({
        'bold': False,   # 字体加粗
        'align': 'left',
        'valign': 'vcenter',
        'border': 2,
        'font_name': u'宋体',
        'font_size': 9
    })
    field = [u'日期', u'证券代码', u'证券名称', u'委托方向', u'委托时间', u'买卖规律', u'委托价格',
             u'前最高点', u'前最低点', u'距前高点', u'距前低点']
    sheet.write_row(0, 0, field, wrapstyle)
    sheet.write_column(1, 0, datelist, otherstyle)
    sheet.write_column(1, 1, codelist, otherstyle)
    sheet.write_column(1, 2, namelist, otherstyle)
    sheet.write_column(1, 3, directionlist, otherstyle)
    sheet.write_column(1, 4, timelist, otherstyle)
    sheet.write_column(1, 5, lawlist, otherstyle)
    sheet.write_column(1, 6, [round(item, 2) for item in pricelist], otherstyle)
    sheet.write_column(1, 7, highlist, otherstyle)
    sheet.write_column(1, 8, lowlist, otherstyle)
    sheet.write_column(1, 9, highprolist, decimalstyle)
    sheet.write_column(1, 10, lowprolist, decimalstyle)
    wb.close()
    return 'Done!'
if __name__ == '__main__':
    print tradeanalysis()
