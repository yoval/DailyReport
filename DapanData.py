# -*- coding: utf-8 -*-
"""
Created on Sat Mar  5 10:14:59 2022

@author: Administrator
"""
import pandas as pd
import glob,os,time,Config

floder = Config.get_floder()
outputfloder = Config.get_outputfloder()
date = time.strftime("%Y-%m-%d", time.localtime())
def loadfiles():
    fileList = glob.glob(os.path.join(floder,'*.xlsx'))
    for fileName in fileList:
        if date in fileName:
            if '服务结果指标' in fileName:
                print('正在读取:%s'%fileName)
                DF_jieguozhibiao = pd.read_excel(fileName,sheet_name=0)
            elif '服务能效指标' in fileName:
                print('正在读取:%s'%fileName)
                DF_nengxiaozhibiao = pd.read_excel(fileName,sheet_name=0)
    return DF_nengxiaozhibiao,DF_jieguozhibiao

DF_nengxiaozhibiao,DF_jieguozhibiao = loadfiles()
DF_nengxiaozhibiao.rename(columns={'进线量':"业务线"},inplace=True)
DF_dapan = pd.concat([DF_jieguozhibiao,DF_nengxiaozhibiao],axis=1)
DF_dapan.rename(columns={'平均排队时长(s)':"平均排队时长"},inplace=True)
DF_dapan.rename(columns={'平均响应时长(s)':"平均响应时长"},inplace=True)
DF_dapan.rename(columns={'平均首次响应时长(s)':"平均首次响应时长"},inplace=True)
DF_dapan.rename(columns={'平均会话时长(s)':"平均会话时长"},inplace=True)
DF_dapan = DF_dapan[['日期','业务线','邀评量','邀评率','参评量','参评率','解决量','解决率','24H解决会话量','24h人工客服首次解决率','48H解决会话量','48h人工客服首次解决率','72H解决会话量','72h人工客服首次解决率','好评量','满意度','中评量','中评率','差评量','差评率','进线量','接起量','接起率','平均排队时长','平均响应时长','平均首次响应时长','平均会话时长','平均人工服务时长','15s首响率','30s首响率','60s首响率','15s响应率','60s响应率']]
DF_dapan_save = DF_dapan.head()
DF_dapan_save.to_csv(outputfloder+date+'-大盘数据源.csv',encoding="utf_8_sig",index=False)
print('已保存')