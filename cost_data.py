# coding:utf-8

import os
import pandas as pd

period = input('请输入统计月份：')
data_path1 = 'E:/data/3-结果数据/6-收入&成本核算/{}/银行与普惠金融服务事业部/个人业务部/'.format(period)
data_path2 = 'E:/data/3-结果数据/6-收入&成本核算/{}/银行与普惠金融服务事业部/银行与普惠金融服务事业部/'.format(period)
name_list1 = os.listdir(data_path1)
name_list2 = os.listdir(data_path2)
['00-01事业部业务条线成本汇总表.xlsx', '00-02事业部本部门及分公司成本汇总表.xlsx', '01账户支付（成本）计价总表.xlsx', '01账户支付（成本）计价总表2.xlsx',
 '02协议支付（成本）计价总表.xlsx', '02协议支付（成本）计价总表2.xlsx', '04-01验证成本一.xlsx', '04-01验证成本一2.xlsx', '04-04验证成本四.xlsx',
 '04-04验证成本四2.xlsx', '04-05验证成本五.xlsx', '04-05验证成本五2.xlsx', '05实名支付手续费（收入）计价总表.xlsx', '05实名支付手续费（收入）计价总表2.xlsx',
 '07-46 收银宝计价总表20210222-0324-事业部-个人.xlsx',
 '08-01 个人事业部相关成本-2021年02月（仅事业部）.xlsx',
 '09金服宝成本计价总表.xlsx', '10垫资成本总表.xlsx', '10垫资成本总表2.xlsx', '12网联网络服务费成本计价总表.xlsx', '12网联网络服务费成本计价总表2.xlsx']


def data_plus(path1, path2):
    for i in os.listdir(path2):
        if i[-7:] == '总表.xlsx':
            if i[-8:] == '汇总表.xlsx':
                pass
            else:
                dataframe1 = pd.read_excel(path1 + i, index_col=0)
                dataframe1.fillna(0, inplace=True)
                dataframe2 = pd.read_excel(path2 + i, index_col=0)
                dataframe2.fillna(0, inplace=True)
        else:
            dataframe3 = pd.read_excel(path1 + i, index_col=0)
            dataframe3.fillna(0, inplace=True)
    return ([dataframe1, dataframe2, dataframe3])
