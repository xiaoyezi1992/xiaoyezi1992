# coding:utf-8

import pandas as pd
import datetime
import numpy as np

# 将各产品明细表按需汇总统计数据

dataPath = 'E:/数据/4-日报表&周报表/日报&周报202010/源数据/'
docDate = input('请输入需打开文件后缀日期（例如20201030）:') # 一般为下载报表日期
readDataDate = input('请输入用户统计数据日期:')
amtDataDate = input('请输入放款统计数据日期：')
orgDate = readDataDate[0:6] + '01' # 用于统计月累计

total = pd.ExcelWriter(dataPath + '产品数据汇总{}.xlsx'.format(readDataDate))

# 助贷用户数据(未完成)
data_user = pd.read_excel((dataPath + '用户报表试验{}.xlsx'.format(docDate)),header=1, usecols=['客户手机号', '申请时间'],index_col=1)
data_user['申请时间'] = pd.to_datetime(data_user['申请时间'])


# 通联钱包数据(已完成)
# wallet_user = pd.read_excel((dataPath + '表1个人会员信息期间汇总报表{}-{}.xls'.format(readDataDate,readDataDate)),
#                           sheet_name='个人会员信息期间汇总报表', header=1,index_col=0,
#                             usecols=['分公司名称', '本期会员数', '新增会员数', '活跃用户数', '当年累计活跃用户数'])
# wallet_user2 = pd.read_excel((dataPath + '表1个人会员信息期间汇总报表{}-{}.xls'.format(orgDate,readDataDate)),
#                           sheet_name='个人会员信息期间汇总报表', header=1,index_col=0,usecols=['分公司名称', '活跃用户数'])
#
# df_wallet = pd.DataFrame({'新增会员数': wallet_user.loc['合计：', '新增会员数'],
# '活跃用户数': wallet_user.loc['合计：', '活跃用户数'],
# '当月累计活跃用户数': wallet_user2.loc['合计：', '活跃用户数'],
# '当年累计活跃用户数': wallet_user.loc['合计：', '当年累计活跃用户数'],
# '注册会员数': wallet_user.loc['合计：', '本期会员数']}, index=[0])
# df_wallet.to_excel(total, '通联钱包')
# total.save()

# 放款数据（pos贷已完成）
# pos_data = pd.read_excel((dataPath + 'posedksq{}.xlsx'.format(docDate)), usecols=['支用起始日', '支用金额'], index_col=1)
# pos_amt = pos_data.groupby(by='支用起始日').sum()

# print('生意金网商贷：')
# print('生意金其他：')
# print('到手花')