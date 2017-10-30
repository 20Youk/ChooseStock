# -*- coding:utf8 -*-
# Author: Youk.Lin
import os
import xlsxwriter
import xlrd
import pandas as pd


def ShowFiles(filepath):
    fileList = os.listdir(filepath)
    absFileList = []
    for item in fileList:
        if os.path.isfile(os.path.join(filepath, item)):
            absFileList.append(os.path.join(filepath, item))
    return absFileList


def Top10(filelist):
    for item in filelist:
        wb = xlrd.open_workbook(item)
        wbName = os.path.split(item)[1]
        sheet = wb.sheet_by_index(0)
        dateList = sheet.col_values(0, start_rowx=1)
        codeList = sheet.col_values(1, start_rowx=1)
        nameList = sheet.col_values(2, start_rowx=1)
        mvList = sheet.col_values(3, start_rowx=1)
        dateFrame = pd.DataFrame({'date': dateList, 'code': codeList, 'name': nameList, 'mv': mvList},
                                 columns=['date', 'code', 'name', 'mv'])
        result = dateFrame.sort_values(by='mv', ascending=False).head(10)


if __name__ == '__main__':
    filePath = 'C:\Users\Youk\Desktop\work'
    print ShowFiles(filePath)
