# coding:utf-8

# 将日报表各产品明细表按需汇总统计数据


import pandas as pd
import datetime


# 确定统计路径和时间
dataPath = 'E:/data/1-原始数据表/产品/'
savePath = 'E:/data/2-数据源表/产品/'
docDate = input('请输入文件下载日期（例如20201030）:')
readDataDate = input('请输入日报表统计日期:')
afterDate = (datetime.datetime.strptime(readDataDate, '%Y%m%d') + datetime.timedelta(days=1)).strftime('%Y%m%d')
beforeDate = (datetime.datetime.strptime(readDataDate, '%Y%m%d') + datetime.timedelta(days=-1)).strftime('%Y%m%d')
lastAmtDate = (datetime.datetime.strptime(readDataDate, '%Y%m%d') + datetime.timedelta(days=-2)).strftime('%Y%m%d')
# 用于生意金已累计放款数据查询
orgDate = input('请输入统计日期的上月末日期:')  # 用于统计助贷用户月累计

lastData = pd.read_excel('E:/data/4-日报表&周报表/日报&周报202010/个人业务事业部日报表_{}.xlsx'.format(beforeDate),
                         sheet_name='Sheet1', header=1,
                         usecols=['区间', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', '月累计', '年累计'])
total = pd.ExcelWriter(savePath + '日报表产品数据{}.xlsx'.format(readDataDate))


# 通联钱包数据
wallet_user = pd.read_excel((dataPath + '表1个人会员信息期间汇总报表_{}_{}.xls'.format(readDataDate,readDataDate)),
                          sheet_name='个人会员信息期间汇总报表', header=1,index_col=0,
                            usecols=['分公司名称', '本期会员数', '新增会员数', '活跃用户数', '当年累计活跃用户数'])
wallet_user2 = pd.read_excel((dataPath + '表1个人会员信息期间汇总报表_{}_{}.xls'.format((readDataDate[0:6] + '01'),
                                                                           readDataDate)),
                             sheet_name='个人会员信息期间汇总报表', header=1,index_col=0,usecols=['分公司名称', '活跃用户数'])
dict_wallet = {'新增用户': int(wallet_user.loc['合计：', '新增会员数'].replace(',', '')),
               '活跃用户': int(wallet_user.loc['合计：', '活跃用户数'].replace(',', ''))}
df_wallet = pd.DataFrame.from_dict(dict_wallet, orient='index', columns=[readDataDate])
df_wallet.index.name = '指标'
df_wallet.loc['活跃用户', '月累计'] = int(wallet_user2.loc['合计：', '活跃用户数'].replace(',', ''))
df_wallet.loc['活跃用户', '年累计'] = int(wallet_user.loc['合计：', '当年累计活跃用户数'].replace(',', ''))
df_wallet.loc['总用户', '年累计'] = int(wallet_user.loc['合计：', '本期会员数'].replace(',', ''))
if readDataDate[-4:] == '0101':
    df_wallet.loc['新增用户', '月累计'] = int(wallet_user.loc['合计：', '新增会员数'].replace(',', ''))
    df_wallet.loc['新增用户', '年累计'] = int(wallet_user.loc['合计：', '新增会员数'].replace(',', ''))
elif readDataDate[-2:] == '01':
    df_wallet.loc['新增用户', '月累计'] = int(wallet_user.loc['合计：', '新增会员数'].replace(',', ''))
    df_wallet.loc['新增用户', '年累计'] = int(lastData.iloc[4, 5]) +\
                                   int(wallet_user.loc['合计：', '新增会员数'].replace(',', ''))
else:
    df_wallet.loc['新增用户', '月累计'] = int(lastData.iloc[4, 4]) +\
                                   int(wallet_user.loc['合计：', '新增会员数'].replace(',', ''))
    df_wallet.loc['新增用户', '年累计'] = int(lastData.iloc[4, 5]) +\
                                   int(wallet_user.loc['合计：', '新增会员数'].replace(',', ''))
msgUser = input('请输入{}短信引流客户数：'.format(readDataDate))
df_wallet.loc['短信引流', readDataDate] = int(msgUser)
df_wallet.loc['短信引流', '月累计'] = int(lastData.iloc[9, 4]) + int(msgUser)
df_wallet.loc['短信引流', '年累计'] = int(lastData.iloc[9, 5]) + int(msgUser)
df_wallet.to_excel(total, '通联钱包用户数据')


# 助贷用户数据
data_user = pd.read_excel((dataPath + '用户报表{}.xlsx'.format(docDate)), header=1, usecols=['客户手机号', '申请时间'])
data_user['申请时间'] = pd.to_datetime(data_user['申请时间'])
data_user.set_index('申请时间', inplace=True)
data_user = pd.Series(data_user['客户手机号'],index=data_user.index)
dict_loan_user = {'新增用户': (len(set(list(data_user[afterDate:]))) - len(set(list(data_user[readDataDate:])))),
                  '活跃用户': len(set(list(data_user[readDataDate])))}
df_loan_user = pd.DataFrame.from_dict(dict_loan_user,orient='index',columns=['数值'])
df_loan_user.index.name = '指标'
df_loan_user.loc['活跃用户', '月累计'] = len(set(list(data_user[afterDate: orgDate])))
df_loan_user.loc['活跃用户', '年累计'] = len(set(list(data_user[afterDate:])))
if readDataDate[-4:] == '0101':
    df_loan_user.loc['新增用户', '月累计'] = (len(set(list(data_user[afterDate:]))) - len(set(list(data_user[readDataDate:]))))
    df_loan_user.loc['新增用户', '年累计'] = (len(set(list(data_user[afterDate:]))) - len(set(list(data_user[readDataDate:]))))
elif readDataDate[-2:] == '01':
    df_loan_user.loc['新增用户', '月累计'] = (len(set(list(data_user[afterDate:]))) - len(set(list(data_user[readDataDate:]))))
    df_loan_user.loc['新增用户', '年累计'] = int(lastData.iloc[7, 5]) +\
                                   (len(set(list(data_user[afterDate:]))) - len(set(list(data_user[readDataDate:]))))
else:
    df_loan_user.loc['新增用户', '月累计'] = int(lastData.iloc[7, 4]) +\
                                   (len(set(list(data_user[afterDate:]))) - len(set(list(data_user[readDataDate:]))))
    df_loan_user.loc['新增用户', '年累计'] = int(lastData.iloc[7, 5]) +\
                                   (len(set(list(data_user[afterDate:]))) - len(set(list(data_user[readDataDate:]))))
df_loan_user.to_excel(total, '助贷用户数据')


# 放款数据
# pos贷
pos_data = pd.read_excel((dataPath + 'posedksq{}.xlsx'.format(docDate)), usecols=['支用起始日', '支用金额'])
pos_data['支用起始日'] = pd.to_datetime(pos_data['支用起始日'])
pos_data.set_index('支用起始日', inplace=True)
pos_data = pd.Series(pos_data['支用金额'],index=pos_data.index)
pos_amt = pos_data[beforeDate].sum()

# 创客贷
ck_data = pd.read_excel((dataPath + 'CKDSJ{}.xlsx'.format(docDate)), header= 1, usecols=['对账日期', '当日放款金额'])
ck_data['对账日期'] = pd.to_datetime(ck_data['对账日期'])
ck_data.set_index('对账日期', inplace=True)
ck_data = pd.Series(ck_data['当日放款金额'],index=ck_data.index)
ck_amt = ck_data[beforeDate].sum()

# 特享贷
tx_data = pd.read_excel((dataPath + 'TXDSJ{}.xlsx'.format(docDate)), header= 1, usecols=['支用时间', '交易金额'])
tx_data['支用时间'] = pd.to_datetime(tx_data['支用时间'], format='%Y%m%d')
tx_data.set_index('支用时间', inplace=True)
tx_data = pd.Series(tx_data['交易金额'],index=tx_data.index)
tx_amt = tx_data[beforeDate].sum()

# 富通贷
ft_data = pd.read_excel((dataPath + '富通贷贷后数据{}.xlsx'.format(docDate)), header= 1, usecols=['支用日期', '支用金额'])
ft_data['支用日期'] = pd.to_datetime(ft_data['支用日期'], format='%Y%m%d')
ft_data.set_index('支用日期', inplace=True)
ft_data = pd.Series(ft_data['支用金额'],index=ft_data.index)
ft_amt = ft_data[beforeDate].sum()

# 通联快贷
tl_data = pd.read_excel((dataPath + '通联快贷贷后数据{}.xlsx'.format(docDate)), header= 1, usecols=['支用日期', '支用金额'])
tl_data['支用日期'] = pd.to_datetime(tl_data['支用日期'])
tl_data.set_index('支用日期', inplace=True)
tl_data = pd.Series(tl_data['支用金额'],index=tl_data.index)
tl_amt = tl_data[beforeDate].sum()

# 生意金
syj_data = pd.read_csv((dataPath + 'tonglian_jigou_{}.csv'.format(beforeDate)), usecols=['AMT'])
syj_data2 = pd.read_csv((dataPath + 'tonglian_jigou_{}.csv'.format(lastAmtDate)), usecols=['AMT'])
syj_amt = int(syj_data.sum() - syj_data2.sum())

# 到手商城
ds_data = pd.read_excel((dataPath + '订单列表{}.xls'.format(beforeDate)), usecols=['订单状态', '订单金额', '期数'])
ds_data.set_index(['订单状态'], inplace=True)
list = ['待发货', '已发货', '备货中']
judge_list = [i in list for i in ds_data.index]
df_ds = ds_data.loc[judge_list]
df_ds.set_index('期数', inplace=True)
ds_amt = int(df_ds.loc[df_ds.index > 0, :].sum())

# 到手现金借款
jk_data = pd.read_excel((dataPath + '借款订单列表{}.xls'.format(beforeDate)), usecols=['订单状态', '借款金额'])
jk_data.set_index(['订单状态'], inplace=True)
list_jk = ['放款中', '分期还款中', '已完成']
judge_list_jk = [j in list_jk for j in jk_data.index]
jk_amt = int(jk_data.loc[judge_list_jk].sum())
totalAmt = (syj_amt + pos_amt + ck_amt + tx_amt + ft_amt + tl_amt + ds_amt + jk_amt)/10000
syjOtherAmt = (pos_amt + ck_amt + tx_amt + ft_amt + tl_amt)/10000
dsTotalAmt = (ds_amt + jk_amt)/10000

dict_loan_amt = {'新增放款（万）': totalAmt,
                 '生意金-网商贷': syj_amt/10000,
                 '生意金-其他': syjOtherAmt,
                 '到手': dsTotalAmt}
df_loan_amt = pd.DataFrame.from_dict(dict_loan_amt, orient='index',columns=['数值'])
df_loan_amt.index.name = '指标'
if beforeDate[-4:] == '0101':
    df_loan_amt.loc['新增放款（万）', '月累计'] = totalAmt
    df_loan_amt.loc['生意金-网商贷', '月累计'] = syj_amt/10000
    df_loan_amt.loc['生意金-其他', '月累计'] = syjOtherAmt
    df_loan_amt.loc['到手', '月累计'] = dsTotalAmt
    df_loan_amt.loc['新增放款（万）', '年累计'] = totalAmt
    df_loan_amt.loc['生意金-网商贷', '年累计'] = syj_amt/10000
    df_loan_amt.loc['生意金-其他', '年累计'] = syjOtherAmt
    df_loan_amt.loc['到手', '年累计'] = dsTotalAmt
elif beforeDate[-2:] == '01':
    df_loan_amt.loc['新增放款（万）', '月累计'] = totalAmt
    df_loan_amt.loc['生意金-网商贷', '月累计'] = syj_amt/10000
    df_loan_amt.loc['生意金-其他', '月累计'] = syjOtherAmt
    df_loan_amt.loc['到手', '月累计'] = dsTotalAmt
    df_loan_amt.loc['新增放款（万）', '年累计'] = lastData.iloc[10, 5] + totalAmt
    df_loan_amt.loc['生意金-网商贷', '年累计'] = lastData.iloc[11, 5] + syj_amt/10000
    df_loan_amt.loc['生意金-其他', '年累计'] = lastData.iloc[12, 5] + syjOtherAmt
    df_loan_amt.loc['到手', '年累计'] = lastData.iloc[13, 5] + dsTotalAmt
else:
    df_loan_amt.loc['新增放款（万）', '月累计'] = lastData.iloc[10, 4] + totalAmt
    df_loan_amt.loc['生意金-网商贷', '月累计'] = lastData.iloc[11, 4] + syj_amt/10000
    df_loan_amt.loc['生意金-其他', '月累计'] = lastData.iloc[12, 4] + syjOtherAmt
    df_loan_amt.loc['到手', '月累计'] = lastData.iloc[13, 4] + dsTotalAmt
    df_loan_amt.loc['新增放款（万）', '年累计'] = lastData.iloc[10, 5] + totalAmt
    df_loan_amt.loc['生意金-网商贷', '年累计'] = lastData.iloc[11, 5] + syj_amt/10000
    df_loan_amt.loc['生意金-其他', '年累计'] = lastData.iloc[12, 5] + syjOtherAmt
    df_loan_amt.loc['到手', '年累计'] = lastData.iloc[13, 5] + dsTotalAmt
df_loan_amt.to_excel(total, '助贷放款')
total.save()

print('完成！' * 10)