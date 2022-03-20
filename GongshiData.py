# -*- coding: utf-8 -*-
"""
Created on Sat Mar  5 13:37:41 2022

@author: Administrator
"""

import pandas as pd
import glob,os,time,datetime,Config

floder = Config.get_floder()
outputfloder = Config.get_outputfloder()
ribaopath = Config.get_ribaopath()
paiban_sheetname = Config.get_paiban_sheetname()
date = time.strftime("%Y-%m-%d", time.localtime())
fileList = glob.glob(os.path.join(floder,'*.xlsx'))

def loadfiles():
    for fileName in fileList:
        if date in fileName:
            if '人员维度' in fileName:
                print('正在读取:%s'%fileName)
                DF_renyuanweidu  = pd.read_excel(fileName,sheet_name=0)
    return DF_renyuanweidu


DF_renyuanweidu = loadfiles()
columns = ['日期','坐席姓名','坐席ID','租户名称','1级节点名称','2级节点名称','3级节点名称','4级节点名称','在线时长','离线时长','小休时长','忙碌时长','培训时长','就餐时长','会议时长','登出时长','外呼时长','BPO','上线时间','员工类型','组别','团队','职场','工时利用率','签入时长','班次','业务线']
df_renyuan = pd.DataFrame(columns=columns)
renyuanweidu_columns = DF_renyuanweidu.columns.values
for column_name in  renyuanweidu_columns:
    if column_name =='统计日期':
        new_column_name = '日期'
    elif column_name == '供应商名称':
        new_column_name = '租户名称'
    elif column_name == '客服id':
        new_column_name = '坐席ID'
    elif column_name == '客服名称':
        new_column_name = '坐席姓名'
    elif column_name == '外呼通话时长':
        new_column_name = '外呼时长'
        
    elif '小时' in column_name:
        new_column_name = column_name.split('-')[0]
        new_column_name = new_column_name.split('(')[0]
    elif column_name in ['接起量','外呼量','外呼接通量','30s首响量','30s平均响应率','完结工单量','平均会话时长','外呼平均通话时长','CPD','付薪时长(不含外呼)']:
        continue
    else:
        new_column_name= column_name
    df_renyuan[new_column_name] = DF_renyuanweidu[column_name]
    
list_BPO = [Config.getBPO(name) for name in df_renyuan.坐席姓名]
list_Name = [name.split('-')[-1] for name in df_renyuan.坐席姓名]
df_renyuan['BPO'] = list_BPO
df_renyuan['坐席姓名'] = list_Name
list_xianlu = []
list_banci = []
df_paiban = pd.read_excel(ribaopath,sheet_name=paiban_sheetname)
for row in df_renyuan.itertuples():
    BPO = row.BPO
    if BPO =='春客':
        Name = row.坐席姓名
        DF_paiban_row = df_paiban[df_paiban['姓名'].isin([Name])]
        try:
            xianlu = list(DF_paiban_row.iloc[0])[-2]
        except:
            xianlu = '未查询到，请更新班表！'   
        try:
            Date = datetime.datetime.strptime(row.日期, "%Y-%m-%d")
            banci = DF_paiban_row[Date].iloc[0]
        except:
            banci = '未检测到，请更新排班！'
    else:
        xianlu = ''
        banci = ''
    list_xianlu.append(xianlu)
    list_banci.append(banci)
del df_paiban
df_renyuan['业务线'] = list_xianlu
df_renyuan['班次'] = list_banci
df_renyuan.to_csv(outputfloder+date+'-工时利用率数据源.csv',encoding="utf_8_sig",index=False)
print('已保存')