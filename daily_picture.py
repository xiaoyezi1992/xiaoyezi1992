# coding:utf-8

import pandas as pd
from matplotlib import pyplot as plt
import datetime

# 数据读取
data_path = 'E:/data/3-结果数据/1-日报表&周报表/日报&周报202010/个人业务事业部日报表_'
date_today = input('请输入日报统计日期：')


def png_data(path, date):
    time_date_today = datetime.datetime.strptime(date, '%Y%m%d')
    # time_date_begin = time_date_today - datetime.timedelta(91)
    period_day = (time_date_today - datetime.datetime.strptime('20210101', '%Y%m%d')).days
    data_dict = {'笔数（万）': [], '金额（亿）': [], '手续费（万）': [], '收益（剔渠道成本/万）': [], '钱包新增用户': [],
                 '钱包活跃用户': [], '助贷活跃用户': [], '新增放款(万)': []}
    data_df = pd.DataFrame.from_dict(data_dict)
    data_df.index.name = '区间'
    for j in range(period_day + 1):
        read_date = (datetime.datetime.strptime('20210101', '%Y%m%d') + datetime.timedelta(j)).strftime('%Y%m%d')
        data = pd.read_excel(path + '{}.xlsx'.format(read_date), header=1, nrows=11)
        data_df.loc[read_date, '笔数（万）'] = data.iloc[0, 5]
        data_df.loc[read_date, '金额（亿）'] = data.iloc[1, 5]
        data_df.loc[read_date, '手续费（万）'] = data.iloc[2, 5]
        data_df.loc[read_date, '收益（剔渠道成本/万）'] = data.iloc[3, 5]
        data_df.loc[read_date, '钱包新增用户'] = data.iloc[4, 5]
        data_df.loc[read_date, '钱包活跃用户'] = data.iloc[5, 5]
        data_df.loc[read_date, '助贷活跃用户'] = data.iloc[8, 5]
        data_df.loc[read_date, '新增放款(万)'] = data.iloc[10, 5]
    return data_df


df = png_data(data_path, date_today)
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


# 传入数组及列名计算该列移动平均值
def avg(dataframe1, col):
    y = []
    for i in range(1, dataframe1.shape[0]):
        sum_data = 0
        avg_data = 0
        j = 0
        while j < i:
            sum_data = sum_data + dataframe1[col][j + 1]
            if j == 0:
                avg_data = sum_data
            else:
                avg_data = sum_data / j
            j += 1
        y.append(avg_data)
    return y


y4_2 = avg(df, '新增放款(万)')


def word_cut(word):
    word = word[-4:]
    return word


x = x_a.map(word_cut)
x_min1 = x_min1_a.map(word_cut)

# 设置图像参数并画图
fig = plt.figure(figsize=(17, 10), dpi=180)
plt.rcParams['font.family'] = 'Microsoft YaHei'  # 正常显示标签
plt.rcParams['axes.unicode_minus'] = False  # 正常显示负号

ax1 = fig.add_subplot(2, 2, 1)
ax1.plot(x, y1_1, color='#84C1FF', label='笔数_万')
ax1.plot(x, y1_2, color='#BBFFBB', label='金额_亿')
ax1.plot(x, y1_3, color='#CE0000', label='手续费_万')
ax1.plot(x, y1_4, color='#FF8000', label='收益_万')
line1 = plt.gca()
line1.spines['right'].set_color('none')
line1.spines['top'].set_color('none')
plt.legend(loc='upper right', fontsize=8)
plt.xticks(x[::7], rotation=-30, fontsize=8)
plt.grid(alpha=0.2)
plt.title('2021TLT支付业务数据变化情况', fontsize=10)

ax2 = fig.add_subplot(2, 2, 2)
ax2.plot(x, y3_1, label='助贷活跃用户', color='#84C1FF')
line2 = plt.gca()
line2.spines['right'].set_color('none')
line2.spines['top'].set_color('none')
plt.legend(loc='upper right', fontsize=8)
plt.xlabel('时间', loc='right', fontsize=8)
plt.xticks(x[::7], rotation=-30, fontsize=8)
plt.grid(alpha=0.2)
plt.title('2021助贷日活变化情况', fontsize=10)

ax3 = fig.add_subplot(2, 2, 3)
ax3.plot(x, y2_1, label='新增用户', color='#84C1FF')
ax3.plot(x, y2_2, label='活跃用户', color='#BBFFBB')
line3 = plt.gca()
line3.spines['right'].set_color('none')
line3.spines['top'].set_color('none')
plt.legend(loc='upper right', fontsize=8)
plt.xticks(x[::7], rotation=-30, fontsize=8)
plt.grid(alpha=0.2)
plt.title('2021C端业务(小通智推)日活变化情况', fontsize=10)

ax4 = fig.add_subplot(2, 2, 4)
ax4.plot(x_min1, y4_1[1:], label='助贷新增放款_万', color='#84C1FF')
ax4.plot(x_min1, y4_2, label='移动平均', color='#CE0000', alpha=0.4, linestyle='-')
line4 = plt.gca()
line4.spines['right'].set_color('none')
line4.spines['top'].set_color('none')
plt.legend(loc='upper right', fontsize=8)
plt.xlabel('时间', loc='right', fontsize=8)
plt.xticks(x[::7], rotation=-30, fontsize=8)
plt.grid(alpha=0.2)
plt.title('2021助贷日放款额变化情况', fontsize=10)

plt.savefig('E:/data/3-结果数据/1-日报表&周报表/日报&周报202010/{}图.png'.format(date_today))
plt.show()