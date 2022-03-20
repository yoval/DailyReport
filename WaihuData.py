# -*- coding: utf-8 -*-
"""
Created on Sat Mar  5 14:37:56 2022

@author: Administrator
"""

import pandas as pd
import glob,os,time,Config


floder = Config.get_floder()
outputfloder = Config.get_outputfloder()
ribaopath = Config.get_ribaopath()
paiban_sheetname = Config.get_paiban_sheetname()
date = time.strftime("%Y-%m-%d", time.localtime())
fileList = glob.glob(os.path.join(floder,'*.xlsx'))

def loadfiles():
    for fileName in fileList:
        if date in fileName:
            if '在线外呼明细' in fileName:
                print('正在读取:%s'%fileName)
                DF_Waihu  = pd.read_excel(fileName,sheet_name=0)
    return DF_Waihu

DF_Waihu = loadfiles()

columns = ['通话日期','员工姓名','在线职场','account_id','外呼量','外呼接起量','外呼接起率','外呼平均振铃时长','外呼平均通话时长','外呼平均话后处理时长','总振铃时长','总通话时长','总话后处理时长','BPO','姓名','职场','组别','上线时间','上线天数','线路']
df_waihu = pd.DataFrame(columns=columns)
waihu_columns = DF_Waihu.columns.values
for column_name in  waihu_columns:
    if column_name =='统计日期':
        new_column_name = '通话日期'
    elif column_name =='员工ID':
        new_column_name = 'account_id'
    elif '(s)' in column_name :
        new_column_name = column_name.split('(')[0]
    elif column_name =='三级节点名称':
        new_column_name = '在线职场'
    elif column_name in ['外呼平均振铃时长(s)','外呼平均通话时长(s)'] :
        continue
    else:
        new_column_name = column_name
    df_waihu[new_column_name] = DF_Waihu[column_name]
    
list_BPO = [Config.getBPO(name) for name in df_waihu.员工姓名]
list_Name = [name.split('-')[-1] for name in df_waihu.员工姓名]
df_waihu['BPO'] = list_BPO

#根据"员工姓名"（七星潍坊在线客服-韩晴晴）匹配线路
list_xianlu = []
df_paiban = pd.read_excel(ribaopath,sheet_name=paiban_sheetname)
for row in df_waihu.itertuples():
    BPO = row.BPO
    if BPO =='春客':
        name = row.员工姓名
        Name = name.split('-')[-1]
        DF_paiban_row = df_paiban[df_paiban['姓名'].isin([Name])]
        try:
            xianlu = list(DF_paiban_row.iloc[0])[-2]
        except:
            xianlu = '未查询到，请更新班表！'        
    else:
        xianlu = ''
    list_xianlu.append(xianlu)
del df_paiban

df_waihu['线路'] = list_xianlu
df_waihu.to_csv(outputfloder + date+'-外呼数据源.csv',encoding="utf_8_sig",index=False)
print('已保存')



