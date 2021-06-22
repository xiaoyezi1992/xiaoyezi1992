# coding:utf-8

import os
import pandas as pd

period = input('请输入数据入账月份（例202104）：')
data_path_gr = 'E:/data/3-结果数据/6-收入&成本核算/{}/银行与普惠金融服务事业部/个人业务部/'.format(period)
data_path_py = 'E:/data/3-结果数据/6-收入&成本核算/{}/银行与普惠金融服务事业部/银行与普惠金融服务事业部/'.format(period)
name_list_gr = os.listdir(data_path_gr)
name_list_py = os.listdir(data_path_py)
all_list = ['00-01事业部业务条线成本汇总表.xlsx', '00-02事业部本部门及分公司成本汇总表.xlsx']
gr_list = ['01账户支付（成本）计价总表.xlsx', '02协议支付（成本）计价总表.xlsx', '04-01验证成本一.xlsx', '04-04验证成本四.xlsx', '04-05验证成本五.xlsx',
           '05实名支付手续费（收入）计价总表.xlsx', '07-46 收银宝计价总表20210222-0324-事业部-个人.xlsx', '08-01 个人事业部相关成本-2021年02月（仅事业部）.xlsx',
           '12网联网络服务费成本计价总表.xlsx']
py_list = ['01账户支付（成本）计价总表.xlsx','10垫资成本总表.xlsx']


def data_plus(path1, path2, list1, list2):
    tx_data1 = pd.read_excel(path1 + '00-01事业部业务条线成本汇总表.xlsx')
    tx_data1.fillna(0, inplace=True)
    zb_data1 = pd.read_excel(path1 + '00-02事业部本部门及分公司成本汇总表.xlsx')
    zb_data1.fillna(0, inplace=True)
    tx_data2 = pd.read_excel(path2 + '00-01事业部业务条线成本汇总表.xlsx')
    tx_data2.fillna(0, inplace=True)
    zb_data2 = pd.read_excel(path2 + '00-02事业部本部门及分公司成本汇总表.xlsx')
    zb_data2.fillna(0, inplace=True)
    tx_data = tx_data1 + tx_data2
    zb_data = zb_data1 + zb_data2

    for i in list1:
        if i.contains('收银宝'):
            syb_book = pd.ExcelFile(path1 + i).sheet_names
            for j in syb_book:
                df = pd.read_excel(path1 + i, sheet_name=j)
                pd.concat([])
        dataframe3 = pd.read_excel(path1 + i, index_col=0)
        dataframe3.fillna(0, inplace=True)

    for i in list2:
        dataframe3 = pd.read_excel(path1 + i, index_col=0)
        dataframe3.fillna(0, inplace=True)
    com_cost_data = pd.concat([com_cost_data, dataFrame3])

    return ([tx_data, zb_data, com_cost_data])
