# -*- coding: utf-8 -*-
"""
Created on Mon Mar  7 16:37:01 2022

@author: Administrator
"""
import configparser, glob,os

conf = configparser.ConfigParser()
conf.read('config.ini', encoding="utf-8-sig")
#获取飞书下载目录
def get_floder():
    floder = conf.get('FilePath','FeishuDownloadPath')
    print('设置的主目录是：%s'%floder)
    return floder
#日报路径
def get_ribaopath():
    ribaopath = conf.get('Ribao','ribaopath')
    print('设置的日报路径为：%s'%ribaopath)
    return ribaopath
#排班表名称
def get_paiban_sheetname():
    paiban_sheetname = conf.get('Ribao','Paiban_sheetname')
    print('配置的排班表名称为：%s'%paiban_sheetname)
    return paiban_sheetname
#延满表名称
def get_yanman_sheetname():
    yanman_sheetname = conf.get('Ribao','Yanman_sheetname')
    print('配置的延满表为：%s'%yanman_sheetname)
    return yanman_sheetname
def get_shengjilvpath():
    sjlpath = conf.get('FilePath','ShengjiPath')
    fileList = glob.glob(os.path.join(sjlpath,'升级率人员达成*.csv'))
    fileList.sort(key=lambda x:os.path.getmtime(x))#按照修改时间排序文件
    fileName = fileList[-1]
    print('检测到的最新升级率文件为:%s'%fileName)
    return fileName
def get_outputfloder():
    outputfloder = conf.get('Out','Outfloder')
    print('配置的输出路径为：%s'%outputfloder)
    return outputfloder
def get_yanmanpath():
    yanmanpath = conf.get('CloudPath','YanmanPath')
    print('配置的延满路径为：%s'%yanmanpath)
    return yanmanpath
#通过创建者名字获得BPO
def getBPO(name):
    if '春客' in name:
        BPO = '春客'
    elif '软通' in name:
        BPO = '软通'
    elif '泰盈' in name:
        BPO = '泰盈'
    elif '博岳' in name:
        BPO = '博岳'
    elif '金慧' in name:
        BPO = '金慧'
    elif '七星' in name:
        BPO = '七星'
    else:
        BPO = ''
    return BPO
#通过工单创建组获得线路
def getxianlu(dqclz):
    if dqclz in ['软通在线','泰盈在线','七星在线','博岳在线','春客在线','金慧在线']:
        xianlu = '大盘'
    elif '1.5' in dqclz:
        xianlu = '1.5'
    elif 'j端' in dqclz or 'J端' in dqclz:
        xianlu = 'J端'
    elif '高价值' in dqclz:
        xianlu = 'C5'
    else:
        xianlu =''
    return xianlu

if __name__ == '__main__':
    get_floder()
    get_ribaopath()
    get_paiban_sheetname()
    get_yanman_sheetname()
    get_outputfloder()
    get_yanmanpath()