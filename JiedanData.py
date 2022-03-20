# -*- coding: utf-8 -*-
"""
#职场1.5线CPD明细
Created on Sat Mar  5 15:19:51 2022

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
            if '一线完结工单量明细-在线' in fileName:
                print('正在读取:%s'%fileName)
                DF_Jiedan  = pd.read_excel(fileName,sheet_name=0)
            elif '职场1.5线CPD明细' in fileName:
                print('正在读取:%s'%fileName)
                DF_YDWCPD  = pd.read_excel(fileName,sheet_name=0)
    return DF_Jiedan,DF_YDWCPD

def loadjiedan():
    for fileName in fileList:
        if date in fileName:
            if '一线完结工单量明细-在线' in fileName:
                print('正在读取:%s'%fileName)
                DF_Jiedan  = pd.read_excel(fileName,sheet_name=0)
    return DF_Jiedan

try:
    DF_Jiedan,DF_YDWCPD = loadfiles()
    DF_YDWCPD = DF_YDWCPD[DF_YDWCPD['当前处理组名称']=="在线-J端1.5线组"]
    DF_YDWCPD.rename(columns={'完结日期':"完结时间"},inplace=True)
    DF_YDWCPD.rename(columns={'CPD':"完结工单量"},inplace=True)
    del DF_YDWCPD['处理人所属公司']
    DF_Jiedan = pd.concat([DF_Jiedan,DF_YDWCPD])
except:
    print('-'*6)
    print('生成结果缺少1.5CPD数据……')
    DF_Jiedan = loadjiedan()

columns = ['日期','业务线划分11.5','当前处理人id','处理者名称','当前处理组名称','完结工单量','姓名','BPO','职场','主管','组别','上线时间','人员属性','线路','出勤','周期','月份']
df_jiedan = pd.DataFrame(columns=columns)
Jiedan_columns = DF_Jiedan.columns.values
for column_name in  DF_Jiedan:
    if column_name =='完结时间':
        new_column_name = '日期'
    else:
        new_column_name = column_name
    df_jiedan[new_column_name] = DF_Jiedan[column_name]        

list_BPO = [Config.getBPO(name) for name in df_jiedan.处理者名称]
list_Name = [str(name).split('-')[-1] for name in df_jiedan.处理者名称]
df_jiedan['BPO'] = list_BPO
df_jiedan['姓名'] = list_Name
list_xianlu = []
list_banci = []
df_paiban = pd.read_excel(ribaopath,sheet_name=paiban_sheetname)
for row in df_jiedan.itertuples():
    BPO = row.BPO
    if BPO =='春客':
        clzmc = row.当前处理组名称
        if clzmc == '春客在线':
            xianlu = '大盘'
        elif '1.5' in clzmc:
            xianlu = '1.5'
        elif '高价值' in clzmc:
            xianlu = 'C5'
        else:
            xianlu = ''
        if xianlu == '1.5':
            Name = row.姓名
            Date = datetime.datetime.strptime(row.日期, "%Y-%m-%d")
            DF_paiban_row = df_paiban[df_paiban['姓名'].isin([Name])]
            try:
                banci = DF_paiban_row[Date].iloc[0]
            except:
                banci = '未查询到班次！'
        else:
            banci = ''
    else:
        xianlu = ''
        banci = ''
    list_xianlu.append(xianlu)
    list_banci.append(banci)
del df_paiban
df_jiedan['线路'] = list_xianlu
df_jiedan['出勤'] = list_banci
df_jiedan.to_csv(outputfloder+date+'-结单数据源.csv',encoding="utf_8_sig",index=False)
print('已保存')