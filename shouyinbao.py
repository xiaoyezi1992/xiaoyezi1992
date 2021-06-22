# _*_coding:utf-8_*_
import pandas as pd
import time

start = time.time()
sub_path = input('请输入月份文件夹：')
total_path = 'E:/data/3-结果数据/6-收入&成本核算/'
data = pd.read_excel(total_path + '{}/收银宝{}.xlsx'.format(sub_path, sub_path), header=1)
data1 = data[data['所属事业部'] == '总部个人事业部']
data1.loc[:, '成本'] = data.loc[:, '交易成本'] + data1.loc[:, '差异成本']
data1.loc[:, '收益'] = data1.loc[:, '已收手续费'] - data1.loc[:, '成本']


# 网关
data_wg = data1.loc[data1.loc[:, '交易类型'].str.contains('网关支付'), :]
total_wg = data_wg.groupby('所属分公司').sum()
total_wg.loc['合计', :] = total_wg.sum()
# 快捷
data2 = data1.drop(data1.loc[data1.loc[:, '交易类型'].str.contains('网关支付'), :].index)
data_kj = data2.loc[data2.loc[:, '交易类型'].str.contains('快捷'), :]
total_kj = data_kj.groupby('所属分公司').sum()
total_kj.loc['合计', :] = total_kj.sum()
# 扫码
data_sm = data2.loc[data2.loc[:, '交易类型'].isin(['QQ钱包支付', '收银套餐购买', '微信支付', '微信退货', '银联扫码', '支付宝支付',
                                               '支付宝退货', '扫码预消费完成', '银联扫码支付']), :]
total_sm = data_sm.groupby('所属分公司').sum()
total_sm.loc['合计', :] = total_sm.sum()
# 收单
data3 = data2.drop(data2.loc[data2.loc[:, '交易类型'].str.contains('快捷'), :].index)
data_sd = data3.drop(data3.loc[data3.loc[:, '交易类型'].isin(['QQ钱包支付', '收银套餐购买', '微信支付', '微信退货', '银联扫码', '支付宝支付',
                                                          '支付宝退货', '扫码预消费完成', '银联扫码支付']), :].index)
total_sd = data_sd.groupby('所属分公司').sum()
total_sd.loc['合计', :] = total_sd.sum()

total_syb = pd.ExcelWriter(total_path + sub_path + '/收银宝汇总{}.xlsx'.format(sub_path))
data_wg.to_excel(total_syb, '网关')
total_wg.to_excel(total_syb, '网关汇总')
data_kj.to_excel(total_syb, '快捷')
total_kj.to_excel(total_syb, '快捷汇总')
data_sm.to_excel(total_syb, '扫码')
total_sm.to_excel(total_syb, '扫码汇总')
data_sd.to_excel(total_syb, '收单')
total_sd.to_excel(total_syb, '收单汇总')
total_syb.save()
end = time.time()
print(end - start)
print('----------\n' * 5)
