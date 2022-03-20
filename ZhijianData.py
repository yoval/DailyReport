# -*- coding: utf-8 -*-
"""
Created on Mon Mar  7 14:00:09 2022

@author: Administrator

内部质检及甲方质检
"""
import pandas as pd
import glob,os,time,Config

floder = Config.get_floder()
ribaopath = Config.get_ribaopath()
paiban_sheetname = Config.get_paiban_sheetname()
outputfloder = Config.get_outputfloder()
fileList = glob.glob(os.path.join(floder,'春客在线日报*.xlsx'))
fileList.sort(key=lambda x:os.path.getmtime(x))#按照修改时间排序文件
fileName = fileList[-1] #最新文件
date = time.strftime("%Y-%m-%d", time.localtime())

# 质检明细 内部质检率数据源  
# sec质检明细  甲方质检数据源  
data = [('质检明细','内部质检率数据源'),
        ('sec质检明细 ','甲方质检数据源')
        ]
for CKSheetName,RibaoSheetName in data:
    def loadfiles():
        print('正在读取:%s'%fileName)
        DF_Neibuzhijian = pd.read_excel(fileName,sheet_name=CKSheetName)
        return DF_Neibuzhijian
    DF_Neibuzhijian = loadfiles()
    Lastupdate = DF_Neibuzhijian.tail(1).日期.iloc[0]
    print('当前春客在线日报文件的最新日期为:%s'%Lastupdate)
    DF_Ribao_neibuzhijian = pd.read_excel(ribaopath,sheet_name=RibaoSheetName).tail(1)
    print('当前日报“内部质检”最新日期为：%s'%DF_Ribao_neibuzhijian.一检质检时间.iloc[0])
    ribao_columns = DF_Ribao_neibuzhijian.columns.values #日报首列
    del DF_Ribao_neibuzhijian
    #新建一个和日报相同的DataFrame
    df_Neibu = pd.DataFrame(columns=ribao_columns)
    LastDF = DF_Neibuzhijian[DF_Neibuzhijian.日期==Lastupdate]
    df_Neibu['一检质检时间'] = LastDF['日期']
    df_Neibu['会话生成时间'] = [i.strftime('%m-%d') for i in LastDF['会话生成时间']]
    df_Neibu['客服姓名'] = LastDF['客服姓名']
    df_Neibu['一检总分'] = LastDF['及格']
    df_paiban = pd.read_excel(ribaopath,sheet_name=paiban_sheetname)
    list_xianlu = []
    df_paiban = pd.read_excel(ribaopath,sheet_name=paiban_sheetname)
    for row in df_Neibu.itertuples():
        Name = row.客服姓名
        DF_paiban_row = df_paiban[df_paiban['姓名'].isin([Name])]
        try:
            xianlu = list(DF_paiban_row.iloc[0])[-2]
        except:
            xianlu = ''
        list_xianlu.append(xianlu)
    del df_paiban
    df_Neibu['业务线'] = list_xianlu
    df_Neibu.to_csv(outputfloder+date+'-'+RibaoSheetName+'.csv',encoding="utf_8_sig",index=False)
print('已保存')