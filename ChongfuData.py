# -*- coding: utf-8 -*-
"""
Created on Sat Mar  5 12:36:16 2022

@author: Administrator
"""

import pandas as pd
import glob,os,time,Config



date = time.strftime("%Y-%m-%d", time.localtime())
#date = '2022-03-13'

floder = Config.get_floder()
outputfloder = Config.get_outputfloder()
fileList = glob.glob(os.path.join(floder,'*.xlsx'))
def loadfiles():
    for fileName in fileList:
        if date in fileName:
            if '在线职场重复进线' in fileName:
                print('正在读取:%s'%fileName)
                DF_Zaixianzhichang = pd.read_excel(fileName,sheet_name=0)
            elif 'j端1.5重复进线' in fileName:
                print('正在读取:%s'%fileName)
                DF_Jduan = pd.read_excel(fileName,sheet_name=0)
            elif '软通高价职场重复进线' in fileName:
                print('正在读取:%s'%fileName)
                DF_Ruantong  = pd.read_excel(fileName,sheet_name=0)
    return DF_Zaixianzhichang,DF_Jduan,DF_Ruantong

DF_Zaixianzhichang,DF_Jduan,DF_Ruantong = loadfiles()
columns = ['日期','当前处理人名称','当前处理组','72h重复进线率','72h产生重复进线工单量','用户主动进线产生工单量','BPO','姓名','职场','组别','主管','上线时间','上线天数','人员性质','周期','月份','线路']
df_Chongfu = pd.DataFrame(columns=columns)
DF_Hebing = pd.concat([DF_Zaixianzhichang,DF_Jduan,DF_Ruantong])
Hebing_columns = DF_Hebing.columns.values
for column_name in  Hebing_columns:
    if column_name == '统计日期':
        new_column_name = '日期'
    elif column_name == '处理人所属公司':
        continue
    else:
        new_column_name= column_name
    df_Chongfu[new_column_name] = DF_Hebing[column_name]

list_BPO = [Config.getBPO(name) for name in df_Chongfu.当前处理人名称]
list_Name = [name.split('-')[-1] for name in df_Chongfu.当前处理人名称]
df_Chongfu['BPO'] = list_BPO
list_xianlu = []
for row in df_Chongfu.itertuples():
    dqclz = row.当前处理组
    xianlu = Config.getxianlu(dqclz)
    list_xianlu.append(xianlu)

df_Chongfu['线路'] = list_xianlu
df_Chongfu = df_Chongfu.sort_values(by='日期',ascending=True) #按日期排序
df_Chongfu.to_csv(outputfloder+date+'-重复进线数据源.csv',encoding="utf_8_sig",index=False)
print('已保存')