# coding:utf-8

# 将日报表支付+个人科技明细表按需汇总统计数据
# 源数据：前一日日报表，明细：支付下载统计日、通联钱包下载统计日及当月累计至统计日、到手下载统计日
# 生意金数据为新统计表，其他助贷产品放款及助贷用户下载当天报表，短信引流数据在统一报表平台下载统计日
# 支付数据直接引用当日已处理通联通数据


import pandas as pd
import datetime
import time

start = time.time()
# 确定统计路径和时间
doc_date = input('请输入文件下载日期（例如20201030）:')
count_date = input('请输入日报表统计日期:')
tlt_path = 'E:/data/1-原始数据表/TLT/每日明细/'
before_date = (datetime.datetime.strptime(count_date, '%Y%m%d') + datetime.timedelta(days=-1)).strftime('%Y%m%d')
last_amt_date = (datetime.datetime.strptime(count_date, '%Y%m%d') + datetime.timedelta(days=-2)).strftime('%Y%m%d')
# 用于生意金已累计放款数据查询
last_month_dt = (datetime.datetime.strptime(count_date[:6] + '01', '%Y%m%d') + datetime.timedelta(days=-1)).strftime(
    '%Y%m%d')  # 上月末，用于统计助贷用户月累计
last_year_dt = (datetime.datetime.strptime((count_date[:4] + '0101'), '%Y%m%d') + datetime.timedelta(days=-1)).strftime(
    '%Y%m%d')  # 上年末，用于统计助贷用户年累计
data_path = 'E:/data/1-原始数据表/产品/'
save_path = 'E:/data/2-数据源表/产品/'
last_data = pd.read_excel('E:/data/3-结果数据/1-日报表&周报表/日报&周报202010/个人业务事业部日报表_{}.xlsx'.format(before_date),
                          sheet_name='Sheet1', header=1, usecols=['区间', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', '月累计',
                                                                  '年累计'])
total = pd.ExcelWriter(save_path + '日报表数据{}.xlsx'.format(count_date))


# 支付数据读取
def get_tlt(path, date):
    tlt_zf = pd.read_excel(path + 'TLT源表汇总_{}.xlsx'.format(date), sheet_name='汇总')
    tlt_zf.set_index('指标', inplace=True)
    return tlt_zf


tlt_data = get_tlt('E:/data/2-数据源表/TLT/', count_date)


# 通联钱包数据读取
def get_wallet_user(path, date, last):
    wallet_user = pd.read_excel((path + '表1个人会员信息期间汇总报表_{}_{}.xls'.format(date, date)),
                                sheet_name='个人会员信息期间汇总报表', header=1, index_col=0,
                                usecols=['分公司名称', '本期会员数', '新增会员数', '活跃用户数', '当年累计活跃用户数'])
    wallet_user2 = pd.read_excel((path + '表1个人会员信息期间汇总报表_{}_{}.xls'.format((date[0:6] + '01'), date)),
                                 sheet_name='个人会员信息期间汇总报表', header=1, index_col=0, usecols=['分公司名称', '活跃用户数'])
    dict_wallet = {'新增用户': int(wallet_user.loc['合计：', '新增会员数'].replace(',', '')),
                   '活跃用户': int(wallet_user.loc['合计：', '活跃用户数'].replace(',', ''))}
    df_wallet = pd.DataFrame.from_dict(dict_wallet, orient='index', columns=[date])
    df_wallet.index.name = '指标'
    df_wallet.loc['活跃用户', '月累计'] = int(wallet_user2.loc['合计：', '活跃用户数'].replace(',', ''))
    df_wallet.loc['活跃用户', '年累计'] = int(wallet_user.loc['合计：', '当年累计活跃用户数'].replace(',', ''))
    df_wallet.loc['总用户', '年累计'] = int(wallet_user.loc['合计：', '本期会员数'].replace(',', ''))
    if date[-4:] == '0101':
        df_wallet.loc['新增用户', '月累计'] = int(wallet_user.loc['合计：', '新增会员数'].replace(',', ''))
        df_wallet.loc['新增用户', '年累计'] = int(wallet_user.loc['合计：', '新增会员数'].replace(',', ''))
    elif date[-2:] == '01':
        df_wallet.loc['新增用户', '月累计'] = int(wallet_user.loc['合计：', '新增会员数'].replace(',', ''))
        df_wallet.loc['新增用户', '年累计'] = int(last.iloc[4, 5]) + int(wallet_user.loc['合计：', '新增会员数'].replace(',', ''))
    else:
        df_wallet.loc['新增用户', '月累计'] = int(last.iloc[4, 4]) + \
                                       int(wallet_user.loc['合计：', '新增会员数'].replace(',', ''))
        df_wallet.loc['新增用户', '年累计'] = int(last.iloc[4, 5]) + int(wallet_user.loc['合计：', '新增会员数'].replace(',', ''))
    return df_wallet


walletUser = get_wallet_user(data_path, count_date, last_data)


# 助贷用户数据读取
def get_loan_user(path, date1, date2, date3, date4, date5):
    data_user = pd.read_excel((path + '用户报表{}.xlsx'.format(date1)), header=1, usecols=['客户手机号', '申请时间'])
    data_user['申请时间'] = data_user['申请时间'].map(lambda x: x[:10])
    data_user = pd.DataFrame(data_user)
    data_user['申请时间'] = pd.to_datetime(data_user['申请时间'], format='%Y-%m-%d')
    data_user.set_index('申请时间', inplace=True)
    dict_loan_user = {'新增用户': len(list(data_user.loc[date2:, '客户手机号'].unique())) - len(
        list(data_user.loc[date5:, '客户手机号'].unique())),
                      '活跃用户': len(list(data_user.loc[date2: date5, '客户手机号'].unique()))}
    df_loan_user = pd.DataFrame.from_dict(dict_loan_user, orient='index', columns=[date2])
    df_loan_user.index.name = '指标'
    df_loan_user.loc['活跃用户', '月累计'] = len(list(data_user.loc[date2: date3, '客户手机号'].unique()))
    df_loan_user.loc['活跃用户', '年累计'] = len(list(data_user.loc[date2: date4, '客户手机号'].unique()))
    if date2[-4:] == '0101':
        df_loan_user.loc['新增用户', '月累计'] = df_loan_user.loc['新增用户', date2]
        df_loan_user.loc['新增用户', '年累计'] = df_loan_user.loc['新增用户', date2]
    elif date2[-2:] == '01':
        df_loan_user.loc['新增用户', '月累计'] = df_loan_user.loc['新增用户', date2]
        df_loan_user.loc['新增用户', '年累计'] = len(list(data_user.loc[date2:, '客户手机号'].unique())) - len(
            list(data_user.loc[date4:, '客户手机号'].unique()))
    else:
        df_loan_user.loc['新增用户', '月累计'] = len(list(data_user.loc[date2:, '客户手机号'].unique())) - len(
            list(data_user.loc[date3:, '客户手机号'].unique()))
        df_loan_user.loc['新增用户', '年累计'] = len(list(data_user.loc[date2:, '客户手机号'].unique())) - len(
            list(data_user.loc[date4:, '客户手机号'].unique()))
    return df_loan_user


loanUser = get_loan_user(data_path, doc_date, count_date, last_month_dt, last_year_dt, before_date)


# 短信引流数据读取
def get_msg_user(path, date, last):
    msg = pd.read_excel((path + '表16会员拓展统计表_{}_{}.xlsx'.format(date, date)), header=1, usecols=['APPID'])
    msg_user = pd.DataFrame(msg)
    dict_msg_user = {'短信引流': msg_user[msg_user['APPID'].isin(['TLA2020', 'TLA2021'])].count()['APPID']}
    df_msg_user = pd.DataFrame.from_dict(dict_msg_user, orient='index', columns=[date])
    df_msg_user.index.name = '指标'
    if date[-4:] == '0101':
        df_msg_user.loc['短信引流', '月累计'] = dict_msg_user['短信引流']
        df_msg_user.loc['短信引流', '年累计'] = dict_msg_user['短信引流']
    elif date[-2:] == '01':
        df_msg_user.loc['短信引流', '月累计'] = dict_msg_user['短信引流']
        df_msg_user.loc['短信引流', '年累计'] = int(last.iloc[9, 5]) + dict_msg_user['短信引流']
    else:
        df_msg_user.loc['短信引流', '月累计'] = int(last.iloc[9, 4]) + dict_msg_user['短信引流']
        df_msg_user.loc['短信引流', '年累计'] = int(last.iloc[9, 5]) + dict_msg_user['短信引流']
    return df_msg_user


msg_users = get_msg_user(data_path, count_date, last_data)


# 放款数据读取
def get_loan_amt(path, date1, date2, date3, last):
    # pos贷
    pos_data = pd.read_excel((path + 'posedksq{}.xlsx'.format(date1)), usecols=['支用起始日', '支用金额'])
    pos_data['支用起始日'] = pd.to_datetime(pos_data['支用起始日'])
    pos_data.set_index('支用起始日', inplace=True)
    pos_data = pd.Series(pos_data['支用金额'], index=pos_data.index)
    pos_amt = pos_data[date2].sum()
    # 创客贷
    ck_data = pd.read_excel((path + 'CKDSJ{}.xlsx'.format(date1)), header=1, usecols=['对账日期', '当日放款金额'])
    ck_data['对账日期'] = pd.to_datetime(ck_data['对账日期'])
    ck_data.set_index('对账日期', inplace=True)
    ck_data = pd.Series(ck_data['当日放款金额'], index=ck_data.index)
    ck_amt = ck_data[date2].sum()
    # 特享贷
    tx_data = pd.read_excel((path + 'TXDSJ{}.xlsx'.format(date1)), header=1, usecols=['支用时间', '交易金额'])
    tx_data['支用时间'] = pd.to_datetime(tx_data['支用时间'], format='%Y%m%d')
    tx_data.set_index('支用时间', inplace=True)
    tx_data = pd.Series(tx_data['交易金额'], index=tx_data.index)
    tx_amt = tx_data[date2].sum()
    # 富通贷
    ft_data = pd.read_excel((path + '富通贷贷后数据{}.xlsx'.format(date1)), header=1, usecols=['支用日期', '支用金额'])
    ft_data['支用日期'] = pd.to_datetime(ft_data['支用日期'], format='%Y%m%d')
    ft_data.set_index('支用日期', inplace=True)
    ft_data = pd.Series(ft_data['支用金额'], index=ft_data.index)
    ft_amt = ft_data[date2].sum()
    # 生意金
    syj_data = pd.read_excel((path + '生意金汇总数据{}.xlsx'.format(date1)), header=1, usecols=['日期', '当日新增支用 金额', '累计支用金额'])
    syj_data['日期'] = pd.to_datetime(syj_data['日期'], format='%Y%m%d')
    syj_data.set_index('日期', inplace=True)
    if syj_data.loc[date2, '当日新增支用 金额'].empty:  # 增加判断当日是否无数据
        syj_amt = 0
    else:
        if int(syj_data.loc[date2, '当日新增支用 金额']) == 0:  # 如当日新增放款为0，新增使用当日减上日累计数
            if syj_data.loc[date3, '当日新增支用 金额'].empty:
                syj_amt = -1  # 手工处理数据
            else:
                syj_amt = int(syj_data.loc[date2, '累计支用金额']) - int(syj_data.loc[date3, '累计支用金额'])
        elif syj_data.loc[date3, '当日新增支用 金额'].empty:
            syj_amt = -1  # 手工处理数据
        else:
            syj_amt = int(syj_data.loc[date2, '当日新增支用 金额'])
    # 汇总
    total_amt = (syj_amt + pos_amt + ck_amt + tx_amt + ft_amt) / 10000
    syj_other_amt = (pos_amt + ck_amt + tx_amt + ft_amt) / 10000
    dict_loan_amt = {'新增放款（万）': total_amt,
                     '生意金-网商贷': syj_amt / 10000,
                     '生意金-其他': syj_other_amt}
    df_loan_amt = pd.DataFrame.from_dict(dict_loan_amt, orient='index', columns=[date2])
    df_loan_amt.index.name = '指标'
    if before_date[-4:] == '0101':
        df_loan_amt.loc['新增放款（万）', '月累计'] = total_amt
        df_loan_amt.loc['生意金-网商贷', '月累计'] = syj_amt / 10000
        df_loan_amt.loc['生意金-其他', '月累计'] = syj_other_amt
        df_loan_amt.loc['新增放款（万）', '年累计'] = total_amt
        df_loan_amt.loc['生意金-网商贷', '年累计'] = syj_amt / 10000
        df_loan_amt.loc['生意金-其他', '年累计'] = syj_other_amt
    elif before_date[-2:] == '01':
        df_loan_amt.loc['新增放款（万）', '月累计'] = total_amt
        df_loan_amt.loc['生意金-网商贷', '月累计'] = syj_amt / 10000
        df_loan_amt.loc['生意金-其他', '月累计'] = syj_other_amt
        df_loan_amt.loc['新增放款（万）', '年累计'] = last.iloc[10, 5] + total_amt
        df_loan_amt.loc['生意金-网商贷', '年累计'] = last.iloc[11, 5] + syj_amt / 10000
        df_loan_amt.loc['生意金-其他', '年累计'] = last.iloc[12, 5] + syj_other_amt
    else:
        df_loan_amt.loc['新增放款（万）', '月累计'] = last.iloc[10, 4] + total_amt
        df_loan_amt.loc['生意金-网商贷', '月累计'] = last.iloc[11, 4] + syj_amt / 10000
        df_loan_amt.loc['生意金-其他', '月累计'] = last.iloc[12, 4] + syj_other_amt
        df_loan_amt.loc['新增放款（万）', '年累计'] = last.iloc[10, 5] + total_amt
        df_loan_amt.loc['生意金-网商贷', '年累计'] = last.iloc[11, 5] + syj_amt / 10000
        df_loan_amt.loc['生意金-其他', '年累计'] = last.iloc[12, 5] + syj_other_amt
    return df_loan_amt


loan_amt = get_loan_amt(data_path, doc_date, before_date, last_amt_date, last_data)
prd = pd.concat([tlt_data, walletUser, loanUser, msg_users, loan_amt])
prd.to_excel(total, '汇总')
total.save()
end = time.time()
print(end - start)
print('----------\n' * 5)
