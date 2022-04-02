# -*- coding: utf-8 -*-
"""
Created on Sat Mar 12 10:52:22 2022

@author: Administrator
春客判断1.5线
用于在线延迟满意度明细
"""
import pandas as pd
import Config,time,glob,os,datetime



def ydwxl(drclzmc):
    if '1.5线' in drclzmc:
        xianlu = '1'
    else:
        xianlu = ''
    return xianlu

def NewOrOld(PanduanDate,Shangxianriqi):
    sxts = (PanduanDate-Shangxianriqi).days #上线天数
    if sxts > 30:
        ryxz = '老人'
    else:
        ryxz = '新人'
    return ryxz
    

outputfloder = Config.get_outputfloder()
date = time.strftime("%Y-%m-%d", time.localtime())
#date = '3-26'
floder = Config.get_floder()
YanmanPath = Config.get_yanmanpath()
print('检测到的最新延满日表文件是：%s'%YanmanPath)
pd_ol_yanman = pd.read_excel(YanmanPath,sheet_name='source')
old_data = pd_ol_yanman.工单编号
print('检测到%s条记录 ✓'%len(old_data))
print('-'*6)
columns = pd_ol_yanman.columns.values
del pd_ol_yanman
fileList = glob.glob(os.path.join(floder,'在线延迟满意明细表*.xlsx')) #延满
filePath = [i for i in fileList if date in i][0] #延满文件
print('检测到的延满文件为：%s'%filePath)
pd_new_yanman = pd.read_excel(filePath,sheet_name=0)
#1.5延满
try:
    fileList_YDW = glob.glob(os.path.join(floder,'1.5延迟满意度明细*.xlsx'))
    filePath_YDW = [i for i in fileList_YDW if date in i][0] #延满文件
    print('检测到的1.5延满文件为:%s'%filePath_YDW)
    pd_YDW_yanman = pd.read_excel(filePath_YDW,sheet_name=0)
#    pd_YDW_yanman = pd_YDW_yanman.loc[(pd_YDW_yanman['当前处理组名称'].str.contains('J端')) | (pd_YDW_yanman['当前处理组名称'].str.contains('山货'))] #筛选
#    pd_YDW_yanman = pd_YDW_yanman.loc[(pd_YDW_yanman['当前处理组名称'].str.contains('山货'))]
    pd_YDW_yanman = pd_YDW_yanman.loc[(pd_YDW_yanman['当前处理组名称']=='在线-J端1.5线组') | (pd_YDW_yanman['当前处理组名称']=='在线-山货1.5线组')]
    pd_new_yanman = pd.concat([pd_new_yanman,pd_YDW_yanman])
except:
    print('无 1.5延迟满意度明细 数据……')
#####
new_data = pd_new_yanman['工单编号']
print('检测到%s条记录'%len(new_data))
print('-'*6)
print('即将删除重复记录')
NewList = list(set(new_data).difference(set(old_data))) #差集
print('删除%s条重复记录，剩余%s条不重复记录！'%(len(new_data)-len(NewList),len(NewList)))
print('-'*6)
New_pd = pd_new_yanman[pd_new_yanman['工单编号'].isin(NewList)]
NameList = [i.split('-')[-1] for i in New_pd.处理者名称]
BPOList = [Config.getBPO(i) for i in New_pd.处理者名称]
ydwxList = [ydwxl(i) for i in New_pd.当前处理组名称]
New_pd.loc[:,'姓名'] = NameList
New_pd.loc[:,'小组'] = ['' for i in range(New_pd.shape[0])]
New_pd.loc[:,'团队'] = ['' for i in range(New_pd.shape[0])]
New_pd.loc[:,'职场'] = ['' for i in range(New_pd.shape[0])]
New_pd.loc[:,'上线日期'] = ['' for i in range(New_pd.shape[0])]
New_pd.loc[:,'员工属性'] = ['' for i in range(New_pd.shape[0])]
New_pd.loc[:,'BPO'] = BPOList
New_pd.loc[:,'是否1.5线'] = ['' for i in range(New_pd.shape[0])]
New_pd.loc[:,'完结时段'] = ['' for i in range(New_pd.shape[0])]
New_pd.loc[:,'业务订单号'] = [str(i) for i in New_pd['业务订单号']]
###########
#匹配排班表
XiaozhuList = []
TuanduiList = []
ZhichangList = []
ShangxianriqiList = []
YuangongshuxingList = []
ydwxList = []
ribaopath = Config.get_ribaopath()
df_huamingce = pd.read_excel(ribaopath,sheet_name='花名册')
#New_pd = New_pd.loc[New_pd['处理者名称'].str.contains('在线')] #剔除热线人员
New_pd = New_pd.loc[~New_pd['处理者名称'].str.contains('CQC')] #剔除含有CQC的行
for row in New_pd.itertuples():
    BPO = row.BPO
    if BPO == '春客':
        try:
            name = row.姓名
            row_huamingce = df_huamingce[df_huamingce['姓名'].isin([name])]
            Xiaozu = row_huamingce.组长.iloc[0]
            Tuandui = row_huamingce.主管.iloc[0]
            Zhichang = row_huamingce.职场.iloc[0]
            try:
                Shangxianriqi = row_huamingce.上线时间.iloc[0] #datetime.datetime
                SXRQ = Shangxianriqi.strftime('%Y/%m/%d')
                PanduanDate = row.完结时间 #str
                PanduanDate = datetime.datetime.strptime(PanduanDate,'%Y-%m-%d') #datetime.datetime
                Yuangongshuxing = NewOrOld(PanduanDate,Shangxianriqi)
            except :
                Yuangongshuxing = '/'
            if row_huamingce.岗位.iloc[0] != '员工':
                Yuangongshuxing = '管理'
            xianlu = ydwxl(row.当前处理组名称) #判断是否1.5线
        except:
            Xiaozu = ''
            Tuandui = ''
            Zhichang = ''
            SXRQ = ''
            Yuangongshuxing = ''
            xianlu = ''    
    else:
        Xiaozu = ''
        Tuandui = ''
        Zhichang = ''
        SXRQ = ''
        Yuangongshuxing = ''
        xianlu = ''
          
    XiaozhuList.append(Xiaozu)
    TuanduiList.append(Tuandui)
    ZhichangList.append(Zhichang)
    ShangxianriqiList.append(SXRQ)
    YuangongshuxingList.append(Yuangongshuxing)
    ydwxList.append(xianlu)

#New_pd.loc[:,'是否1.5线'] = [int(i) for i in ydwxList]
New_pd.loc[:,'是否1.5线'] =  ydwxList
New_pd.loc[:,'小组'] = XiaozhuList
New_pd.loc[:,'团队'] = TuanduiList
New_pd.loc[:,'职场'] =  ZhichangList
New_pd.loc[:,'上线日期'] =  ShangxianriqiList
New_pd.loc[:,'员工属性'] = YuangongshuxingList
New_pd.loc[:,'完结时间'] 
del df_huamingce

New_pd.to_excel(outputfloder + date+'-YanMan_Table.xlsx',sheet_name = 'sheet1',encoding="utf_8_sig",index=False)
del New_pd
print('已保存')