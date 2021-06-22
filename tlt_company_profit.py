# coding:utf-8
import pandas as pd
import datetime


# 按分公司汇总并三七分润，直营分15%，特殊调整中信百信
def com_total(path, date):
    df_detail = pd.read_excel(path + '{}.xlsx'.format(date),
                              sheet_name='明细', usecols=['商户号', '收入所属方', '笔数', '金额', '手续费', '收益'])
    company_total = df_detail.groupby('收入所属方').sum()
    company_total = company_total[['笔数', '金额', '手续费', '收益']]
    company_ini = company_total * 0.7
    company_ini.loc['直营', :] = company_ini.loc['直营', :] / 0.7 * 0.85
    num_total = df_detail.groupby('商户号').sum()
    num_total = num_total[['笔数', '金额', '手续费', '收益']]
    if 200100000019927 in num_total.index:
        if 200100000023767 in num_total.index:
            adjust_special = num_total.loc[200100000019927, :] * 0.35 + num_total.loc[200100000023767, :] * 0.35
        else:
            adjust_special = num_total.loc[200100000019927, :] * 0.35
    else:
        if 200100000023767 in num_total.index:
            adjust_special = num_total.loc[200100000023767, :] * 0.35
        else:
            adjust_special = 0
    company_ini.loc['湖北分公司', :] = company_ini.loc['湖北分公司', :] - adjust_special
    company_ini.loc[:, '手续费'] = company_ini.loc[:, '手续费'] / 1.06
    company_ini.loc[:, '收益'] = company_ini.loc[:, '收益'] / 1.06
    company_res = company_ini / 10000
    company_res = company_res.reset_index()

    def com_sim(company):
        company = company.replace('分公司', '')
        return company
    company_res.loc[:, '收入所属方'] = company_res.loc[:, '收入所属方'].map(com_sim)
    company_res = company_res.sort_values(by='手续费', ascending=False)
    company_res.loc[:, '收益率'] = company_res.loc[:, '收益'] / company_res.loc[:, '手续费']
    company_res.loc[:, '收入排名'] = company_res.loc[:, '手续费'].rank(method='first', ascending=False)
    return company_res


source_path = 'E:/data/2-数据源表/TLT/TLT源表汇总_'
week_date = input('请输入需统计周度数据的起始日期（例如20210412-20210418）：')
week_company = com_total(source_path, week_date)


def compare_last(last_path, date):
    last_time_start = (datetime.datetime.strptime(date[:8], '%Y%m%d') - datetime.timedelta(7)).strftime('%Y%m%d')
    last_time_end = (datetime.datetime.strptime(date[-8:], '%Y%m%d') - datetime.timedelta(7)).strftime('%Y%m%d')
    last_data = pd.read_excel(last_path + '{}-{}.xlsx'.format(last_time_start, last_time_end), sheet_name='Sheet1',
                              usecols=['分公司', '本年累计收入', '本周收入', '本周收入排名'])
    fill_data = last_data[['分公司', '本年累计收入', '本周收入', '本周收入排名']]
    fill_data.loc[:, '上周收入'] = fill_data.loc[:, '本周收入']
    fill_data.loc[:, '上周收入排名'] = fill_data.loc[:, '本周收入排名']
    fill_data.loc[:, '上周本年累计收入'] = fill_data.loc[:, '本年累计收入']
    fill_data = fill_data[['分公司', '上周本年累计收入', '上周收入', '上周收入排名']]
    return fill_data


res_path = 'E:/data/3-结果数据/2-企业微信号/公众号数据'
compare_company = compare_last(res_path, week_date)

week_company_all = pd.merge(week_company, compare_company, left_on='收入所属方', right_on='分公司')
week_company_all.loc[:, '本年累计收入'] = week_company_all.loc[:, '上周本年累计收入'] + week_company_all.loc[:, '手续费']
week_company_all.loc[:, '收入环比变化'] = week_company_all.loc[:, '手续费'] / week_company_all.loc[:, '上周收入'] - 1
week_company_all.loc[:, '收入排名变化'] = week_company_all.loc[:, '上周收入排名'] - week_company_all.loc[:, '收入排名']
week_company_all = week_company_all[['收入所属方', '本年累计收入', '手续费', '上周收入', '收入环比变化', '上周收入排名', '收入排名',
                                     '收入排名变化', '收益率', '笔数', '金额']]

week_com_result = pd.ExcelWriter('E:/data/2-数据源表/TLT/分润分公司' + '{}.xlsx'.format(week_date))
week_company.to_excel(week_com_result, '本周分公司已分润')
week_company_all.to_excel(week_com_result, '微信公众号数据')
week_com_result.save()

print('----------\n' * 5)
