# -*- coding: utf-8 -*-
"""
Created on Fri Mar  4 19:21:40 2022

@author: fuwen

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
            if '人效数据' in fileName:
                print('正在读取:%s'%fileName)
                DF_renxiaoshuju = pd.read_excel(fileName,sheet_name=0)
            elif '人员维度' in fileName and date in fileName:
                print('正在读取:%s'%fileName)
                DF_renyuanweidu = pd.read_excel(fileName,sheet_name=0)
            elif 'IM30s响应率' in fileName and date in fileName:
                print('正在读取:%s'%fileName)
                DF_xiangyinglv = pd.read_excel(fileName,sheet_name=0)
    return DF_renxiaoshuju ,DF_renyuanweidu ,DF_xiangyinglv

DF_renxiaoshuju ,DF_renyuanweidu ,DF_xiangyinglv = loadfiles()
DF_renxiaoshuju.fillna(0, inplace=True)
DF_renyuanweidu.fillna(0, inplace=True)
DF_xiangyinglv.fillna(0, inplace=True)
list_BPO = []
list_Name = []
list_zxysc = [] #总响应时长
list_zxyxxs=[] #总响应消息数
list_ssmxyl = [] #30秒响应量
list_gyjql = []
columns = ['日期','公司名称','是否外包','员工唯一姓名','职场标识','工号','平均首次响应时长','平均响应时长','平均排队时长','平均会话时长','30s首响率','30s响应率','24h人工客服首次解决率','解决率','接起率','满意度','中评度','差评度','接起量','邀评量','参评量','5、4星评价数','平均分','3星评价数','1，2星评价数','总首次响应时长','30秒内首响会话数','总会话总时长','总响应时长','总响应消息量','30s-响应量','出勤','CPD','上线时间','上线天数','人员性质','组别','管理','职场','班次','周期','月份','业务线','工时','BPO']
df_renxiao = pd.DataFrame(columns=columns)
sy_columns = ['日期','公司名称','是否外包','员工唯一姓名','平均首次响应时长(s)','平均响应时长(s)','平均排队时长(s)','平均会话时长(s)','30s首响率','30s响应率','24h人工客服首次解决率','解决率','接起率','满意度','中评度','差评度','接起量','邀评量','参评量','满意量','平均分']

for sy_col in sy_columns:
    if sy_col == '满意量':
        sy_col_new = '5、4星评价数'
    elif sy_col =='线路':
        sy_col_new ='业务线'
    else :
        sy_col_new = sy_col
    sy_col_new = sy_col_new.replace(r'(s)', '')
    df_renxiao[sy_col_new] = DF_renxiaoshuju[sy_col]

for row in df_renxiao.itertuples():
    name = row.员工唯一姓名
    BPO = Config.getBPO(name)
    Name = name.split('-')[-1]
    DF_renxiaoshuju_row = DF_renxiaoshuju[DF_renxiaoshuju['员工唯一姓名'].isin([name])]
    DF_renyuanweidu_row = DF_renyuanweidu[DF_renyuanweidu['客服名称'].isin([name])]
    DF_xiangyinglv_row = DF_xiangyinglv[DF_xiangyinglv['员工唯一姓名'].isin([name])]
    try:
        gyjql = DF_renyuanweidu_row.接起量 
        gyjql = gyjql.iloc[0]
    except:
        print('人员维度未查询到：%s。即将查询人效数据……'%Name)
        try:
            gyjql = DF_renxiaoshuju_row.接起量 
            gyjql = gyjql.iloc[0]
        
        except:
            gyjql = DF_xiangyinglv_row.接起量
            gyjql = gyjql.iloc[0]
    list_gyjql.append(gyjql) 
    zxysc = DF_xiangyinglv_row.总响应时长.iloc[0]
    zxyxxs = DF_xiangyinglv_row.总响应消息数.iloc[0]
    ssmxyl = DF_xiangyinglv_row.总响应消息数*DF_xiangyinglv_row['30s响应率'] #30秒响应量
    ssmxyl = ssmxyl.iloc[0]
    list_BPO.append(BPO)
    list_Name.append(Name)
    list_zxysc.append(zxysc)
    list_zxyxxs.append(zxyxxs)
    list_ssmxyl.append(ssmxyl)
list_ssmxyl = [round(i) for i in list_ssmxyl ]  #三十秒响应量取整
df_renxiao['BPO'] = list_BPO
df_renxiao['员工唯一姓名'] = list_Name
df_renxiao['总响应时长'] = list_zxysc
df_renxiao['总响应消息量'] = list_zxyxxs
df_renxiao['30s-响应量'] = list_ssmxyl
df_renxiao['接起量'] = list_gyjql
list_xianlu = []
df_paiban = pd.read_excel(ribaopath,sheet_name=paiban_sheetname)
for row in df_renxiao.itertuples():
    BPO = row.BPO
    if BPO =='春客':
        Name = row.员工唯一姓名
        DF_paiban_row = df_paiban[df_paiban['姓名'].isin([Name])]
        try:
            xianlu = list(DF_paiban_row.iloc[0])[-2]
        except:
            xianlu = '未查询到，请更新班表！'   
    else:
        xianlu = ''
    list_xianlu.append(xianlu)
del df_paiban
df_renxiao['业务线'] = list_xianlu
df_renxiao.to_csv(outputfloder+date+'-人效日数据源.csv',encoding="utf_8_sig",index=False)
print('已保存')