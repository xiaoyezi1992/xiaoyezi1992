# coding:utf-8

import os
import pandas as pd


sysPath = 'C:/Users/admin/Desktop/新建文件夹/'
dfs = []
for name in os.listdir(sysPath):
    data = pd.read_excel(sysPath + name)
    dfs.append(data)
df = pd.concat(dfs)
savePath = pd.ExcelWriter('C:/Users/admin/Desktop/汇总_新建文件夹.xlsx')
df.to_excel(savePath)
savePath.save()

