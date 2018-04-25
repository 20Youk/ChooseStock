# -*- coding:utf8 -*-
# Author : Youk
# Description : 将好大夫官网的医生信息，按照病例数和投票数整理成数据表格
import xlrd
import pandas as pd
import numpy as np
import datetime

wb_1 = xlrd.open_workbook(r'C:\Users\Administrator\Desktop\HDF\data20180309.xlsx')
sheet = wb_1.sheet_by_index(0)
nameList = sheet.col_values(0, start_rowx=1)
numList_1 = sheet.col_values(6, start_rowx=1)
numList_2 = sheet.col_values(7, start_rowx=1)
xmDataFrame = pd.DataFrame({})
tpDataFrame = pd.DataFrame({})
for ii in range(0, len(nameList)):
    xmDict = {}
    tpDict = {}
    # 整理临床经验列表
    if u'\u7968' in numList_1[ii]:
        df_xm = pd.DataFrame(index=[nameList[ii]], data=[np.nan], columns=[u'无'])
        xmDataFrame = xmDataFrame.append(df_xm)
    else:
        xmList = numList_1[ii].split(',')
        for jj in range(0, len(xmList)):
            oneXmList = xmList[jj].split('(')
            xmDict[oneXmList[0].strip()] = int(oneXmList[1].strip()[:-2])
        xmDataFrame = xmDataFrame.append(pd.DataFrame(xmDict, index=[nameList[ii]]))
    # 整理患者投票列表
    if u'\u65e0' in numList_2[ii] or u'\u6ca1' in numList_2[ii]:
        tp_xm = pd.DataFrame(index=[nameList[ii]], data=[np.nan], columns=[u'无'])
        tpDataFrame = tpDataFrame.append(tp_xm)
    else:
        tpList = numList_2[ii].split(',')
        for kk in range(0, len(tpList) - 1):
            oneTpList = tpList[kk].split('(')
            tpDict[oneTpList[0].strip()] = int(oneTpList[1].strip()[:-2])
        tpDataFrame = tpDataFrame.append(pd.DataFrame(tpDict, index=[nameList[ii]]))
xmDataFrame = xmDataFrame.drop(columns=u'\u65e0').fillna(value=0)
xmDataFrame.astype(int)
tpDataFrame = tpDataFrame.drop(columns=u'\u65e0').fillna(value=0)
tpDataFrame.astype(int)
today = datetime.datetime.now().strftime('%Y%m%d')
write = pd.ExcelWriter(r'C:\Users\Administrator\Desktop\HDF\result_%s.xlsx' % today)
xmDataFrame.to_excel(write, u'临床经验', index=True)
tpDataFrame.to_excel(write, u'投票', index=True)
print 'Done!'
