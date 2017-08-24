# -*- coding:utf8 -*-
import xlrd
import xlsxwriter


def readexcel():
    sheet = wb.sheet_by_index(0)
    returnlist = sheet.col_values(2, start_rowx=1)
    datelist = sheet.col_values(0, start_rowx=1)
    sheetname = sheet.name
    return returnlist, datelist, sheetname

if __name__ == '__main__':
    try:
        wb = xlrd.open_workbook(r'..\excel\test11.xlsx')
        returnList, dateList, sheetName = readexcel()
        wb1 = xlsxwriter.Workbook('..\excel\Test.xlsx')
        sheet1 = wb1.add_worksheet(sheetName)
        dateStyle = wb1.add_format({
            'bold': False,   # 字体加粗
            'align': 'center',
            'valign': 'vcenter',
            'num_format': 'yyyy/m/d'
        })
        decimalStyle = wb1.add_format({
            'bold': False,   # 字体加粗
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '0.0000_ ;[red]-0.0000'
        })
        wrapStyle = wb1.add_format()
        wrapStyle.set_text_wrap()
        sheet1.set_row(0, cell_format= wrapStyle)
        dateList.insert(0, u'交易日期中文测试')
        sheet1.write_column('A1', dateList)
        returnList.insert(0, u'模拟组合每日收益率')
        sheet1.write_column('B1', returnList)
        sheet1.set_column('A:A', 11, dateStyle)
        sheet1.set_column('B:B', 7, decimalStyle)
        wb1.close()
    except IOError, e:
        print '执行失败\n', e
    else:
        print 'Done!!!'

















































































































































