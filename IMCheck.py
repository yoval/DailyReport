# -*- coding: utf-8 -*-
"""
Created on Sun Mar 20 16:38:39 2022
#导出数据
@author: Administrator
IM大盘
"""
import pandas as pd
import glob,os,time,Config

date = time.strftime("%Y-%m-%d", time.localtime())

floder = Config.get_floder()
outputfloder = Config.get_outputfloder()
fileList = glob.glob(os.path.join(floder,'[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].xlsx')) #13位数字
fileList.sort(key=lambda x:os.path.getmtime(x))#按照修改时间排序文件
fileName = fileList[-1] #最新文件
print('检索到的IM文件为：%s'%fileName)
DF_OutPut  = pd.read_excel(fileName,sheet_name=0)
nameList = [i.split('-')[-1] for i in DF_OutPut.客服姓名]

















