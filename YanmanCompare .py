# -*- coding: utf-8 -*-
"""
Created on Mon Mar 28 10:33:41 2022

@author: F-

对比延满表与延满数据源
"""
import pandas as pd
import Config
#日报延满
outputfloder = Config.get_outputfloder()
ribaopath = Config.get_ribaopath()
pd_ribao_yanman = pd.read_excel(ribaopath,sheet_name='延迟满意度数据源')
ribao_data = pd_ribao_yanman.工单编号
print('日报延满表共有%s条数据'%len(ribao_data))
#延满专表
YanmanPath = Config.get_yanmanpath()
pd_zhuanbiao = pd.read_excel(YanmanPath,sheet_name='source')
yanman_data = pd_zhuanbiao.工单编号
print('延满专项表共有%s条数据'%len(yanman_data))
s= len(yanman_data) - len(ribao_data)
print('两者表相差数据为：%s'%abs(s))
print('*'*6)
NewList = list(set(yanman_data).difference(set(ribao_data))) #差集
if len(NewList) ==0:
    print('延满表数据均在日报表之上√')
    NewList = list(set(ribao_data).difference(set(yanman_data)))
    if len(NewList) ==0:
        print('日报表数据均在延满表之上√')
pd_new_yanman = pd_zhuanbiao[pd_zhuanbiao['工单编号'].isin(NewList)]
pd_new_yanman.to_excel(outputfloder +'YanMan_Table_相差.xlsx',sheet_name = 'sheet1',encoding="utf_8_sig",index=False)
del pd_ribao_yanman
del pd_zhuanbiao