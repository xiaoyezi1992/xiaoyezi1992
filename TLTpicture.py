# coding:utf-8

import pandas as pd
from matplotlib import pyplot as plt
import datetime

dataPath = 'E:/data/4-日报表&周报表/日报&周报202010/个人业务事业部日报表_'
date_today = input('请输入日报统计日期：')
time_date_today = datetime.datetime.strptime(date_today, '%Y%m%d')
time_date_begin = time_date_today - datetime.timedelta(45)
data_dict = {'交易笔数（万）': [], '交易金额（亿）': [], '手续费（万）': [], '收益（剔除渠道成本/万）': []}
df = pd.DataFrame.from_dict(data_dict)
df.index.name = '区间'
for i in range(46):
    read_date = (time_date_begin + datetime.timedelta(i)).strftime('%Y%m%d')
    data = pd.read_excel(dataPath + '{}.xlsx'.format(read_date), header=1, nrows=4)
    df.loc[read_date, '交易笔数（万）'] = data.iloc[0, 5]
    df.loc[read_date, '交易金额（亿）'] = data.iloc[1, 5]
    df.loc[read_date, '手续费（万）'] = data.iloc[2, 5]
    df.loc[read_date, '收益（剔除渠道成本/万）'] = data.iloc[3, 5]

fig = plt.figure()
ax = fig.add_subplot(1, 1, 1)
ax.plot(df['交易笔数（万）'], label='交易笔数')
ax.plot(df['交易金额（亿）'], label='交易金额')
ax.plot(df['手续费（万）'], label='手续费')
ax.plot(df['收益（剔除渠道成本/万）'], label='收益')
ax.legend()
plt.xticks(df.index, rotation=45)
plt.show()

