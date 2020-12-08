# coding:utf-8

import pandas as pd
# 将周报表各产品明细表按需汇总统计数据

# 确定统计路径和时间
dataPath = 'E:/data/1-原始数据表/产品/'
savePath = 'E:/data/2-数据源表/产品/'
beginDate = input('请输入周报表起始日期:')
readDataDate = input('请输入周报表截止日期:')
total = pd.ExcelWriter(savePath + '通联钱包数据汇总{}-{}.xlsx'.format(beginDate, readDataDate))


# 通联钱包数据
wallet_user = pd.read_excel((dataPath + '表1个人会员信息期间汇总报表_{}_{}.xls'.format(beginDate, readDataDate)),
                            sheet_name='个人会员信息期间汇总报表', header=1,index_col=0,
                            usecols=['分公司名称', '本期会员数', '新增会员数', '活跃用户数', '当年累计活跃用户数'])
wallet_user2 = pd.read_excel((dataPath + '表1个人会员信息期间汇总报表_{}_{}.xls'.format((readDataDate[0:6] + '01'), readDataDate)),
                             sheet_name='个人会员信息期间汇总报表', header=1,index_col=0,usecols=['分公司名称', '活跃用户数'])

dict_wallet = {'本周新增会员数': wallet_user.loc['合计：', '新增会员数'],
               '本周活跃用户数': wallet_user.loc['合计：', '活跃用户数'],
               '当月累计活跃用户数': wallet_user2.loc['合计：', '活跃用户数'],
               '当年累计活跃用户数': wallet_user.loc['合计：', '当年累计活跃用户数'],
               '注册会员数': wallet_user.loc['合计：', '本期会员数']}
df_wallet = pd.DataFrame.from_dict(dict_wallet, orient='index',columns=['数值'])
df_wallet.to_excel(total, '通联钱包')
total.save()