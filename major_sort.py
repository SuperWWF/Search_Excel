#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Aug 28 17:55:19 2019
需要自己新建Done和Raw文件夹
@author: zhiyuli
"""
import pandas as pd
import numpy as np
import time
import re
import os

def getListFiles(path):
    ret = [] 
    for root, dirs, files in os.walk(path):  
        for filespath in files: 
            ret.append(os.path.join(root,filespath)) 
    return ret
def find_index (file_name_in, file_name_out):
    #  读取专业目录excel文件
    data_source = pd.read_excel(file_name_in, sheet_name = 'Sheet1')
    # 读取专业目录中用到的列
    data_source = data_source.loc[:,['专业门类','专业类','专业名称']]
    # 读取目标文件夹中的Excel，等待填写
    data_output = pd.read_excel(file_name_out, sheet_name = 'Sheet1')
    # 从0开始循环读取填写文件
    for number in range (0,data_output.shape[0],1):
        # 读取目标文件中的专业名称
        search_key = str(data_output.loc[number,'专业名称'])
        # 去除专业名称中的空格等干扰项
        search_key = re.sub('\（.*?\）', '', search_key)
        # 在专业库中查找专业名称的位置
        index = np.where(data_source == search_key)
        # 判断是否找到
        if len(index[0]) != 0:
            #print (index[0][0],index[1][0])
            # 如果找到，读取此专业名称对应的专业类和专业门类
            Row = index[0][0]
            Col = index[1][0]
            # 读取专业类
            Major_Men = np.array(data_source)[Row,0]
            # 读取专业门类
            Major_Lei = np.array(data_source)[Row,1]
            # 打印出专业类和专业门类观察
            print(number,search_key,Major_Men,Major_Lei)
            # 把专业类和专业门类填写到目标文件中
            data_output.loc[number,'专业门类'] = Major_Men
            data_output.loc[number,'专业类'] = Major_Lei
    # 保存更新后的目标文件到Done文件夹中
    save = './Done/' + file_name_out.lstrip('./RAW/')
    data_output.to_excel(save,index = False)      

# 专业目录读取
file_name_in = '本科专业知识目录库-----新(1).xlsx'
# 目标文件夹读取
file_name_out = './RAW'
# 读取需要填写的文件夹中的文件
file_list = getListFiles (file_name_out)
# 逐个对文件夹进行读取，填写好的文件存放在/Done文件夹下
for each in file_list:
    print(each)
    # 填写文件函数，输入专业目录和要填写的文件
    find_index (file_name_in, each)


