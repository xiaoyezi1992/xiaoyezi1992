# _*_coding:utf-8_*_


import pandas as pd

total_path = 'E:/data/11-收入&成本核算/'
sub_path = input('请输入月份文件夹：')
doc_name = input('请输入文件名称：')
data = pd.read_excel(total_path + sub_path + '/' + doc_name + '.xlsx', header=1)
data1 = data[data['所属事业部'] == '总部个人事业部']
data1.loc[:, '成本'] = data1.loc[:, '交易成本'] + data1.loc[:, '差异成本']
data1.loc[:, '收益'] = data1.loc[:, '已收手续费'] - data1.loc[:, '成本']


def sli(dt):
    return str(dt)[:3]


data_ITS = data1[data1['交易渠道'].map(sli) == 'ITS']
data2 = data1[~(data1['交易渠道'].map(sli) == 'ITS')]
data_kj = data2[data2['交易类型'].str.contains('快捷')]
data_sm = data2[data2['交易类型'].isin(['QQ钱包支付', '收银套餐购买', '微信支付', '微信退货', '银联扫码', '支付宝支付',
                                    '支付宝退货', '扫码预消费完成'])]
data3 = data2[~(data2['交易类型'].str.contains('快捷'))]
data_sd = data3[~(data3['交易类型'].isin(['QQ钱包支付', '收银套餐购买', '微信支付', '微信退货', '银联扫码', '支付宝支付',
                                      '支付宝退货', '扫码预消费完成']))]
total_syb = pd.ExcelWriter(total_path + sub_path + '收银宝汇总.xlsx')
data_ITS.to_excel(total_syb, '网关')
data_kj.to_excel(total_syb, '快捷')
data_sm.to_excel(total_syb, '扫码')
data_sd.to_excel(total_syb, '收单')
total_syb.save()
