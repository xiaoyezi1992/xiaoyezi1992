# coding:utf-8

import pandas as pd
# 将周报表各产品明细表按需汇总统计数据

# 确定统计路径和时间
dataPath = 'E:/数据/1-原始数据表/产品/'
savePath = 'E:/数据/2-数据源表/产品/'
docDate = input('请输入文件下载日期（例如20201030）:')
beginDate = input('请输入周报表起始日期:')
beforeBeginDate = input('周报表起始日期前一天：')
readDataDate = input('请输入周报表截止日期:')
afterDate = input('周报表报表截止日期的后一天：')
orgDate = input('上月末日期:')  # 用于统计月累计
total = pd.ExcelWriter(savePath + '周报表产品数据汇总{}-{}.xlsx'.format(beginDate, readDataDate))

# 助贷用户数据
data_user = pd.read_excel((dataPath + '用户报表{}.xlsx'.format(docDate)), header=1, usecols=['客户手机号', '申请时间'])
data_user['申请时间'] = pd.to_datetime(data_user['申请时间'])
data_user.set_index('申请时间', inplace=True)
data_user = pd.Series(data_user['客户手机号'],index=data_user.index)
dict_loan_user = {'新增用户数': (len(set(list(data_user[afterDate:]))) - len(set(list(data_user[beginDate:])))),
                  '本周活跃用户': len(set(list(data_user[afterDate: beforeBeginDate]))),
                  '当月累计活跃用户': len(set(list(data_user[afterDate: orgDate]))),
                  '累计用户数': len(set(list(data_user[afterDate:])))}
df_loan_user = pd.DataFrame.from_dict(dict_loan_user, orient='index', columns=['数值'])
df_loan_user.to_excel(total, '助贷用户')


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


# 放款数据
# pos贷
pos_data = pd.read_excel((dataPath + 'posedksq{}.xlsx'.format(docDate)), usecols=['支用起始日', '支用金额'])
pos_data['支用起始日'] = pd.to_datetime(pos_data['支用起始日'])
pos_data.set_index('支用起始日', inplace=True)
pos_data = pd.Series(pos_data['支用金额'],index=pos_data.index)
pos_amt = pos_data[beginDate: readDataDate].sum()


# 创客贷
ck_data = pd.read_excel((dataPath + 'CKDSJ{}.xlsx'.format(docDate)), header= 1, usecols=['对账日期', '当日放款金额'])
ck_data['对账日期'] = pd.to_datetime(ck_data['对账日期'])
ck_data.set_index('对账日期', inplace=True)
ck_data = pd.Series(ck_data['当日放款金额'],index=ck_data.index)
ck_amt = ck_data[beginDate: readDataDate].sum()


# 特享贷
tx_data = pd.read_excel((dataPath + 'TXDSJ{}.xlsx'.format(docDate)), header= 1, usecols=['支用时间', '交易金额'])
tx_data['支用时间'] = pd.to_datetime(tx_data['支用时间'], format='%Y%m%d')
tx_data.set_index('支用时间', inplace=True)
tx_data = pd.Series(tx_data['交易金额'],index=tx_data.index)
tx_amt = tx_data[beginDate: readDataDate].sum()


# 富通贷
ft_data = pd.read_excel((dataPath + '富通贷贷后数据{}.xlsx'.format(docDate)), header= 1, usecols=['支用日期', '支用金额'])
ft_data['支用日期'] = pd.to_datetime(ft_data['支用日期'], format='%Y%m%d')
ft_data.set_index('支用日期', inplace=True)
ft_data = pd.Series(ft_data['支用金额'],index=ft_data.index)
ft_amt = ft_data[readDataDate: beforeBeginDate].sum()

# 通联快贷
tl_data = pd.read_excel((dataPath + '通联快贷贷后数据{}.xlsx'.format(docDate)), header= 1, usecols=['支用日期', '支用金额'])
tl_data['支用日期'] = pd.to_datetime(tl_data['支用日期'])
tl_data.set_index('支用日期', inplace=True)
tl_data = pd.Series(tl_data['支用金额'],index=tl_data.index)
tl_amt = tl_data[readDataDate: beforeBeginDate].sum()

# 生意金
syj_data = pd.read_csv((dataPath + 'tonglian_jigou_{}.csv'.format(readDataDate)), usecols=['AMT'])
syj_data2 = pd.read_csv((dataPath + 'tonglian_jigou_{}.csv'.format(beforeBeginDate)), usecols=['AMT'])
syj_amt = int(syj_data.sum() - syj_data2.sum())

# 到手商城
ds_data = pd.read_excel((dataPath + '订单列表{}-{}.xls'.format(beginDate, readDataDate)), usecols=['订单状态', '订单金额', '期数'])
ds_data.set_index(['订单状态'], inplace=True)
list = ['待发货', '已发货', '备货中']
judge_list = []
for i in ds_data.index:
    judge_list.append(i in list)
df_ds = ds_data.loc[judge_list]
df_ds.set_index('期数', inplace=True)
ds_amt = int(df_ds.loc[df_ds.index > 0, :].sum())

# 到手现金借款
jk_data = pd.read_excel((dataPath + '借款订单列表{}-{}.xls'.format(beginDate, readDataDate)), usecols=['订单状态', '借款金额'])
jk_data.set_index(['订单状态'], inplace=True)
list_jk = ['放款中', '分期还款中', '已完成']
judge_list_jk = []
for j in jk_data.index:
    judge_list_jk.append(j in list_jk)
jk_amt = int(jk_data.loc[judge_list_jk].sum())

dict_loan_amt = {'放款金额': (syj_amt + pos_amt + ck_amt + tx_amt + ft_amt + tl_amt + ds_amt + jk_amt)/10000,
                 '生意金-网商贷': syj_amt / 10000,
                 '生意金-其他': (pos_amt + ck_amt + tx_amt + ft_amt + tl_amt) / 10000,
                 '到手': (ds_amt + jk_amt) / 10000}
df_loan_amt = pd.DataFrame.from_dict(dict_loan_amt, orient='index',columns=['数值'])
df_loan_amt.to_excel(total, '助贷放款')
total.save()