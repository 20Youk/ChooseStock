# -*- coding:utf8 -*-
import xlrd
import xlsxwriter


def readexcel():
    sheet = wb.sheet_by_index(0)
    returnlist = sheet.col_values(1, start_rowx=1)
    datelist = sheet.col_values(0, start_rowx=1)
    sheetname = sheet.name
    return returnlist, datelist, sheetname

if __name__ == '__main__':
    try:
        wb = xlrd.open_workbook(r'..\..\excel\test11.xlsx')
        returnList, dateList, sheetName = readexcel()
        wb1 = xlsxwriter.Workbook('..\..\excel\Test.xlsx')
        sheet1 = wb1.add_worksheet(sheetName)
        # 数字格式
        decimalStyle = wb1.add_format({
            'bold': False,   # 字体加粗
            'align': 'left',
            'valign': 'vcenter',
            'num_format': '0.0000_ ;[red]-0.0000',
            'border': 2,
            'font_name': u'宋体',
            'font_size': 9
        })
        # 日期格式
        newStyle = wb1.add_format()
        newStyle.set_border(2)
        newStyle.set_font_size(9)
        newStyle.set_font_name(u'宋体')
        newStyle.set_num_format('yyyy-mm-dd')
        newStyle.set_align('left')      # 左对齐
        newStyle.set_align('vcenter')   # 垂直居中
        # 标题格式
        wrapStyle = wb1.add_format()
        wrapStyle.set_text_wrap()  # 自动换行
        wrapStyle.set_pattern(1)    # 填充整个单元格
        wrapStyle.set_fg_color('#5C5C5C')   # 单元格底色（RGB颜色代码表）
        wrapStyle.set_font_name(u'宋体')  # 字体
        wrapStyle.set_border(2)     # 设置边框（线条类型）
        wrapStyle.set_font_size(9)      # 字体大小
        sheet1.write_row(0, 0, [u'交易日期中文测试', u'模拟组合每日收益率'], wrapStyle)
        sheet1.write_column('A2', dateList, newStyle)
        sheet1.write_column('B2', returnList, decimalStyle)
        sheet1.set_column('A:A', 11)
        sheet1.set_column('B:B', 7)
        wb1.close()
    except IOError, e:
        print '执行失败\n', e
    else:
        print 'Done!!!'

















































































































































