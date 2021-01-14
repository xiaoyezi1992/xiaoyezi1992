# coding:utf-8

import pandas as pd
import datetime
# 将周报表各产品明细表按需汇总统计数据
# 源数据：通联钱包下载统计对应周度日期及当月累计至截止日期，到手下载统计对应周度日期，生意金需要当天及开始日前一天放款数据表
# 其他助贷产品放款及助贷用户下载当天报表

# 确定统计路径和时间
dataPath = 'E:/data/1-原始数据表/产品/'
savePath = 'E:/data/2-数据源表/产品/'
docDate = input('请输入文件下载日期（例如20201030）:')
beginDate = input('请输入周报表起始日期:')
beforeBeginDate = (datetime.datetime.strptime(beginDate, '%Y%m%d') + datetime.timedelta(days=-1)).strftime('%Y%m%d')
readDataDate = input('请输入周报表截止日期:')
afterDate = (datetime.datetime.strptime(readDataDate, '%Y%m%d') + datetime.timedelta(days=1)).strftime('%Y%m%d')
orgDate = (datetime.datetime.strptime(readDataDate[:6] + '01', '%Y%m%d') + datetime.timedelta(days=-1))\
    .strftime('%Y%m%d')  # 上月末日期，用于统计月累计
last_year_date = (datetime.datetime.strptime(readDataDate[:4] + '0101', '%Y%m%d') + datetime.timedelta(days=-1))\
    .strftime('%Y%m%d')
total = pd.ExcelWriter(savePath + '周报表通联钱包{}-{}.xlsx'.format(beginDate, readDataDate))


# 助贷用户数据
def get_loan_user(path, date1, date2, date3, date4, date5, date6):
    data_user = pd.read_excel((path + '用户报表{}.xlsx'.format(date1)), header=1, usecols=['客户手机号', '申请时间'])
    data_user = pd.DataFrame(data_user)
    data_user['申请时间'] = pd.to_datetime(data_user['申请时间'].map(lambda x: x[:10]))
    data_user.set_index('申请时间', inplace=True)
    dict_loan_user = {'新增用户数': len(list(data_user.loc[date2:, '客户手机号'].unique())) - len(
        list(data_user.loc[date3:, '客户手机号'].unique()))}
    df_loan_user = pd.DataFrame.from_dict(dict_loan_user, orient='index', columns=['{}-{}'.format(date5, date2)])
    df_loan_user.loc['新增用户数', '月累计'] = len(list(data_user.loc[date2:, '客户手机号'].unique())) - len(
        list(data_user.loc[date4:, '客户手机号'].unique()))
    df_loan_user.loc['新增用户数', '年累计'] = len(list(data_user.loc[date2:, '客户手机号'].unique())) - len(
        list(data_user.loc[date6:, '客户手机号'].unique()))
    return df_loan_user


loan_user = get_loan_user(dataPath, docDate, readDataDate, beforeBeginDate, orgDate, beginDate, last_year_date)


# 通联钱包数据
def get_wallet_user(path, date1, date2):
    wallet_data1 = pd.read_excel((path + '表1个人会员信息期间汇总报表_{}_{}.xls'.format(date1, date2)),
                                 sheet_name='个人会员信息期间汇总报表', header=1, index_col=0, usecols=
                                 ['分公司名称', '本期会员数', '新增会员数', '活跃用户数', '当年累计活跃用户数'])
    wallet_data2 = pd.read_excel((path + '表1个人会员信息期间汇总报表_{}_{}.xls'.format((date2[0:6] + '01'), date2)),
                                 sheet_name='个人会员信息期间汇总报表', header=1, index_col=0, usecols=['分公司名称',
                                                                                            '新增会员数', '活跃用户数'])

    dict_wallet = {'新增用户': wallet_data1.loc['合计：', '新增会员数'],
                   '活跃用户': wallet_data1.loc['合计：', '活跃用户数']}
    df_wallet = pd.DataFrame.from_dict(dict_wallet, orient='index',columns=['{}-{}'.format(date1, date2)])
    df_wallet.loc['新增用户', '月累计'] = wallet_data2.loc['合计：', '新增会员数']
    year_data = pd.read_excel('E:/data/4-日报表&周报表/日报&周报202010/个人业务事业部日报表_{}.xlsx'.format(date2),
                              sheet_name='Sheet1', header=1, usecols=['区间', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4',
                                                                      '月累计', '年累计'])
    df_wallet.loc['新增用户', '年累计'] = year_data.iloc[4, 5]
    df_wallet.loc['活跃用户', '月累计'] = wallet_data2.loc['合计：', '活跃用户数']
    df_wallet.loc['活跃用户', '年累计'] = wallet_data1.loc['合计：', '当年累计活跃用户数']
    return df_wallet


wallet_user = get_wallet_user(dataPath, beginDate, readDataDate)


# 放款数据
def get_loan_amt(path, date1, date2, date3, date4, date5):
    # pos贷
    pos_data = pd.read_excel((path + 'posedksq{}.xlsx'.format(date1)), usecols=['支用起始日', '支用金额'])
    pos_data['支用起始日'] = pd.to_datetime(pos_data['支用起始日'], format='%Y-%m-%d')
    pos_data.set_index('支用起始日', inplace=True)
    pos_data = pd.Series(pos_data['支用金额'], index=pos_data.index)
    pos_amt = pos_data[date2: date4].sum()

    # 创客贷
    ck_data = pd.read_excel((path + 'CKDSJ{}.xlsx'.format(date1)), header=1, usecols=['对账日期', '当日放款金额'])
    ck_data['对账日期'] = pd.to_datetime(ck_data['对账日期'], format='%Y-%m-%d')
    ck_data.set_index('对账日期', inplace=True)
    ck_data = pd.Series(ck_data['当日放款金额'], index=ck_data.index)
    ck_amt = ck_data[date3: date5].sum()

    # 特享贷
    tx_data = pd.read_excel((path + 'TXDSJ{}.xlsx'.format(date1)), header=1, usecols=['支用时间', '交易金额'])
    tx_data['支用时间'] = pd.to_datetime(tx_data['支用时间'], format='%Y%m%d')
    tx_data.set_index('支用时间', inplace=True)
    tx_data = pd.Series(tx_data['交易金额'],index=tx_data.index)
    tx_data.sort_values()
    tx_amt = tx_data[date3: date5].sum()

    # 富通贷
    ft_data = pd.read_excel((path + '富通贷贷后数据{}.xlsx'.format(date1)), header=1, usecols=['支用日期', '支用金额'])
    ft_data['支用日期'] = pd.to_datetime(ft_data['支用日期'], format='%Y%m%d')
    ft_data.set_index('支用日期', inplace=True)
    ft_data = pd.Series(ft_data['支用金额'], index=ft_data.index)
    ft_amt = ft_data[date2: date4].sum()

    # 通联快贷
    tl_data = pd.read_excel((path + '通联快贷贷后数据{}.xlsx'.format(date1)), header=1, usecols=['支用日期', '支用金额'])
    tl_data['支用日期'] = pd.to_datetime(tl_data['支用日期'])
    tl_data.set_index('支用日期', inplace=True)
    tl_data = pd.Series(tl_data['支用金额'],index=tl_data.index)
    tl_amt = tl_data[date2: date4].sum()

    # 生意金
    syj_data = pd.read_csv((path + 'tonglian_jigou_{}.csv'.format(date2)), usecols=['AMT'])
    syj_data2 = pd.read_csv((path + 'tonglian_jigou_{}.csv'.format(date4)), usecols=['AMT'])
    syj_amt = int(syj_data.sum() - syj_data2.sum())

    # 到手商城
    ds_data = pd.read_excel((path + '订单列表{}-{}.xls'.format(date3, date2)), usecols=['订单状态', '订单金额', '期数'])
    ds_data.set_index(['订单状态'], inplace=True)
    list_ds = ['待发货', '已发货', '备货中']
    judge_list = [i in list_ds for i in ds_data.index]
    df_ds = ds_data.loc[judge_list]
    df_ds.set_index('期数', inplace=True)
    ds_amt = int(df_ds.loc[df_ds.index > 0, :].sum())

    # 到手现金借款
    jk_data = pd.read_excel((path + '借款订单列表{}-{}.xls'.format(date3, date2)), usecols=['订单状态', '借款金额'])
    jk_data.set_index(['订单状态'], inplace=True)
    list_jk = ['放款中', '分期还款中', '已完成']
    judge_list_jk = [j in list_jk for j in jk_data.index]
    jk_amt = int(jk_data.loc[judge_list_jk].sum())

    dict_loan_amt = {'生意金': (syj_amt + pos_amt + ck_amt + tx_amt + ft_amt + tl_amt) / 10000,
                     '到手': (ds_amt + jk_amt) / 10000}
    df_loan_amt = pd.DataFrame.from_dict(dict_loan_amt, orient='index',columns=['{}-{}'.format(date3, date2)])
    year_data = pd.read_excel('E:/data/4-日报表&周报表/日报&周报202010/个人业务事业部日报表_{}.xlsx'.format(date5),
                              sheet_name='Sheet1', header=1, usecols=['区间', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4',
                                                                      '月累计', '年累计'])
    df_loan_amt.loc['生意金', '月累计'] = year_data.iloc[11, 4] + year_data.iloc[12, 4]
    df_loan_amt.loc['生意金', '年累计'] = year_data.iloc[11, 5] + year_data.iloc[12, 5]
    df_loan_amt.loc['到手', '月累计'] = year_data.iloc[13, 4]
    df_loan_amt.loc['到手', '年累计'] = year_data.iloc[13, 5]
    return df_loan_amt


loan_amt = get_loan_amt(dataPath, docDate, readDataDate, beginDate, beforeBeginDate, afterDate)

save_data = pd.concat([wallet_user, loan_user, loan_amt])
save_data.to_excel(total, '周数据汇总')
total.save()

print('数据汇总完成！' * 10)