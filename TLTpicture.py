# coding:utf-8

import pandas as pd
from matplotlib import pyplot as plt
import datetime

dataPath = 'E:/data/4-日报表&周报表/日报&周报202010/个人业务事业部日报表_'
date_today = input('请输入日报统计日期：')
time_date_today = datetime.datetime.strptime(date_today, '%Y%m%d')
time_date_begin = time_date_today + datetime.timedelta(days=-31)
data_dict = {'交易笔数（万）': [], '交易金额（亿）': [], '手续费（万）': [], '收益（剔除渠道成本/万）': []}
df = pd.DataFrame.from_dict(data_dict, orient='index')
df.index.name = '区间'
for i in range(32):
    read_date = (time_date_begin + datetime.timedelta(days=i)).strftime('%Y%m%d')
    data = pd.read_excel(dataPath + '{}.xlsx'.format(read_date), header=1, nrows=4)
    data_cut = data.iloc[:, [1, 5]]
    data_cut.reindex(data_cut['区间'])
    df = df.merge(data_cut, on='区间')
df = df.set_index('区间')

x = df.columns
y1 = df.loc['交易笔数（万）', :]
y2 = df.loc['交易金额（亿）', :]
y3 = df.loc['手续费（万）', :]
y4 = df.loc['收益（剔除渠道成本/万）', :]
plt.plot(x, y1)
plt.plot(x, y2)
plt.plot(x, y3)
plt.plot(x, y4)
plt.xticks(df.columns, [(time_date_begin + datetime.timedelta(days=i)).strftime('%Y-%m-%d') for i in range(32)])
plt.legend(loc="upper right")
plt.show()

