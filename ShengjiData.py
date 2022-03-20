# -*- coding: utf-8 -*-
"""
Created on Sat Mar  5 15:33:22 2022

@author: Administrator

"""
import pandas as pd
import Config,time

outputfloder = Config.get_outputfloder()
filePath = Config.get_shengjilvpath()
date = time.strftime("%Y-%m-%d", time.localtime())
DF_Shenji = pd.read_csv(filePath, encoding = "utf-8")
DF_Shenji = DF_Shenji.loc[DF_Shenji['工单创建组名称'].str.contains('在线')]
DF_Shenji = DF_Shenji.drop_duplicates()

#待生成表格
columns =['创建日期','创建者名字','工单创建组','升级工单量','新增工单量','姓名','组长','主管','职场','上线时间','上线天数','人员性质','周期','月份','BPO','业务线']
df_sj = pd.DataFrame(columns=columns)

list_xianlu = []
for row in DF_Shenji.itertuples():
    dqclz = row.工单创建组名称
    xianlu = Config.getxianlu(dqclz)
    list_xianlu.append(xianlu)

df_sj['创建者名字'] = DF_Shenji['创建者名字']
df_sj['升级工单量'] = DF_Shenji['升级工单量_new']
df_sj['新增工单量'] = DF_Shenji['工单编号']
df_sj['工单创建组'] = DF_Shenji['工单创建组名称']
df_sj['创建日期'] = DF_Shenji['创建日期']
df_sj['业务线'] = list_xianlu
list_BPO = [Config.getBPO(name) for name in df_sj.创建者名字]
df_sj['BPO'] = list_BPO
df_sj.to_csv(outputfloder +date+ '-升级率数据源.csv',encoding="utf_8_sig",index=False)
print('已保存')