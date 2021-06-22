# _*_coding:utf-8_*_
import pandas as pd


def get_data(path, date):
    source_data = pd.read_excel(path + 'P2P平台资金监测表_{}.xlsx'.format(date),
                                sheet_name='明细表', usecols=['商户名称', '当日入金', '当日入金\n与T-30日入金比', '当日出金',
                                                           '当日出金\n与T-30日出金比', '平台涉及\n备付金（元）'])
    source_data.set_index('商户名称', inplace=True)
    source_data = source_data.loc[['北京玖富普惠信息技术有限公司', '上海证大投资咨询有限公司'], :]
    source_data = source_data[['平台涉及\n备付金（元）', '当日入金', '当日出金', '当日入金\n与T-30日入金比',  '当日出金\n与T-30日出金比']]
    source_data.loc[:, '出金金额占备付金比重'] = source_data.loc[:, '当日出金'] / source_data.loc[:, '平台涉及\n备付金（元）']
    source_data.loc['合计', :] = source_data.sum()
    source_data.loc['合计', '当日入金\n与T-30日入金比'] = '-'
    source_data.loc['合计', '当日出金\n与T-30日出金比'] = '-'
    source_data.loc['合计', '出金金额占备付金比重'] = '-'
    source_data = source_data[['平台涉及\n备付金（元）', '当日入金', '当日出金', '出金金额占备付金比重', '当日入金\n与T-30日入金比',  '当日出金\n与T-30日出金比']]
    return source_data


path1 = 'E:/data/1-原始数据表/P2P/'
path2 = 'E:/data/2-数据源表/P2P/'
date1 = input('请输入统计日期:')
P2P_data = get_data(path1, date1)
P2P_save = pd.ExcelWriter(path2 + 'P2P日报{}.xlsx'.format(date1))
P2P_data.to_excel(P2P_save)
P2P_save.save()

print('-----\n' * 5)