# -*- coding: utf-8 -*-
"""
Created on Sun Mar  6 23:00:03 2022

@author: fuwen

通过云端文件提取延满
"""

import pandas as pd
import time,Config

floder = Config.get_floder()
outputfloder = Config.get_outputfloder()
ribaopath = Config.get_ribaopath()
print(ribaopath)
yanman_sheetname = Config.get_yanman_sheetname()
df_ribaoyanman = pd.read_excel(ribaopath,sheet_name=yanman_sheetname).tail(1)
last_gdbh = df_ribaoyanman.工单编号.iloc[0]   #最后一行的工单编号
fileName = Config.get_yanmanpath()
date = time.strftime("%Y-%m-%d", time.localtime())
def loadfiles():
    print('正在读取:%s'%fileName)
    DF_Yanman = pd.read_excel(fileName,sheet_name=0)
    return DF_Yanman
DF_Yanman = loadfiles()
DF_Yanman_row = DF_Yanman[DF_Yanman['工单编号'].isin([last_gdbh])]
last_row_number = DF_Yanman_row.index[0]+1
DF_Yanman = DF_Yanman.iloc[last_row_number:,0:-2]
list_yewu = []
for row in DF_Yanman.itertuples():
    dqclz = row.当前处理组名称
    yewu = Config.getxianlu(dqclz)
    list_yewu.append(yewu)

DF_Yanman['上线天数'] = ['' for i in range(DF_Yanman.shape[0])]
DF_Yanman['周期'] = ['' for i in range(DF_Yanman.shape[0])]
DF_Yanman['月份'] = ['' for i in range(DF_Yanman.shape[0])]
DF_Yanman['业务线'] = list_yewu
DF_Yanman.to_excel(outputfloder + date+'-延满数据源.xlsx',sheet_name = 'sheet1',encoding="utf_8_sig",index=False)
print('已保存')