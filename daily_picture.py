# coding:utf-8

import pandas as pd
from matplotlib import pyplot as plt
import datetime

# 数据读取
data_path = 'E:/data/3-结果数据/1-日报表&周报表/日报&周报202010/个人业务事业部日报表_'
date_today = input('请输入日报统计日期：')
time_date_today = datetime.datetime.strptime(date_today, '%Y%m%d')
# time_date_begin = time_date_today - datetime.timedelta(91)
period_day = (time_date_today - datetime.datetime.strptime('20210101', '%Y%m%d')).days
data_dict = {'笔数（万）': [], '金额（亿）': [], '手续费（万）': [], '收益（剔渠道成本/万）': [], '钱包新增用户': [],
             '钱包活跃用户': [], '助贷活跃用户': [], '新增放款(万)': []}
df = pd.DataFrame.from_dict(data_dict)
df.index.name = '区间'
for i in range(period_day + 1):
    read_date = (datetime.datetime.strptime('20210101', '%Y%m%d') + datetime.timedelta(i)).strftime('%Y%m%d')
    data = pd.read_excel(data_path + '{}.xlsx'.format(read_date), header=1, nrows=11)
    df.loc[read_date, '笔数（万）'] = data.iloc[0, 5]
    df.loc[read_date, '金额（亿）'] = data.iloc[1, 5]
    df.loc[read_date, '手续费（万）'] = data.iloc[2, 5]
    df.loc[read_date, '收益（剔渠道成本/万）'] = data.iloc[3, 5]
    df.loc[read_date, '钱包新增用户'] = data.iloc[4, 5]
    df.loc[read_date, '钱包活跃用户'] = data.iloc[5, 5]
    df.loc[read_date, '助贷活跃用户'] = data.iloc[8, 5]
    df.loc[read_date, '新增放款(万)'] = data.iloc[10, 5]
# 设置画图数据
x_a = df.index
x_min1_a = x_a[:-1]
y1_1 = df['笔数（万）']
y1_2 = df['金额（亿）']
y1_3 = df['手续费（万）']
y1_4 = df['收益（剔渠道成本/万）']
y2_1 = df['钱包新增用户']
y2_2 = df['钱包活跃用户']
y3_1 = df['助贷活跃用户']
y4_1 = df['新增放款(万)']
y4_2 = []
for i in range(1, df.shape[0]):
    sum_data = 0
    avg_data = 0
    j = 0
    while j < i:
        sum_data = sum_data + df['新增放款(万)'][j + 1]
        if j == 0:
            avg_data = sum_data
        else:
            avg_data = sum_data / j
        j += 1
    y4_2.append(avg_data)


def word_cut(word):
    word = word[-4:]
    return word


x = x_a.map(word_cut)
x_min1 = x_min1_a.map(word_cut)

# 设置图像参数并画图
fig = plt.figure(figsize=(18, 10), dpi=100)
plt.rcParams['font.family'] = 'Microsoft YaHei'  # 正常显示标签
plt.rcParams['axes.unicode_minus'] = False  # 正常显示负号

ax1 = fig.add_subplot(2, 2, 1)
ax1.plot(x, y1_1, color='#84C1FF', label='笔数_万')
ax1.plot(x, y1_2, color='#BBFFBB', label='金额_亿')
ax1.plot(x, y1_3, color='#CE0000', linestyle='-.', label='手续费_万')
ax1.plot(x, y1_4, color='#FF8000', linestyle='-.', label='收益_万')
line1 = plt.gca()
line1.spines['right'].set_color('none')
line1.spines['top'].set_color('none')
plt.legend(loc='upper right')
plt.xticks(x[::7], rotation=-60, fontsize=8)
plt.grid(alpha=0.2)
plt.title('TLT支付业务折线图', fontsize=10)

ax2 = fig.add_subplot(2, 2, 2)
ax2.plot(x, y3_1, label='助贷活跃用户', color='#84C1FF')
line2 = plt.gca()
line2.spines['right'].set_color('none')
line2.spines['top'].set_color('none')
plt.legend(loc='upper right')
plt.xlabel('时间', loc='right')
plt.xticks(x[::7], rotation=-60, fontsize=8)
plt.grid(alpha=0.2)

ax3 = fig.add_subplot(2, 2, 3)
ax3.plot(x, y2_1, label='钱包新增用户', color='#84C1FF')
ax3.plot(x, y2_2, label='钱包活跃用户', color='#BBFFBB')
line3 = plt.gca()
line3.spines['right'].set_color('none')
line3.spines['top'].set_color('none')
plt.legend(loc='upper right')
plt.xticks(x[::7], rotation=-60, fontsize=8)
plt.grid(alpha=0.2)

ax4 = fig.add_subplot(2, 2, 4)
ax4.plot(x_min1, y4_1[1:], label='助贷新增放款_万', color='#84C1FF')
ax4.plot(x_min1, y4_2, label='移动平均', color='red', alpha=0.4, linestyle='-')
line4 = plt.gca()
line4.spines['right'].set_color('none')
line4.spines['top'].set_color('none')
plt.legend(loc='upper right')
plt.xlabel('时间', loc='right')
plt.xticks(x[::7], rotation=-60, fontsize=8)
plt.grid(alpha=0.2)

plt.savefig('E:/data/3-结果数据/1-日报表&周报表/日报&周报202010/{}图.jpg'.format(date_today))
plt.show()
