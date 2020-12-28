# _*_coding:utf-8_*_

# 筛选交易类型,验证剔除快捷协议签约申请交易类型的交易明细，取部分分析用字段明细
# 增加笔数、金额、产品、收益、项目标签、商户简称，并修改联营商户收入所属方
# 二级行业为卡中心、交易类型为转账扣款的交易金额、笔数归零；


import pandas as pd
import datetime


# 交易期间选择函数
def period_choose():
    daily_path = 'E:/data/1-原始数据表/TLT/每日明细/'
    month_path = 'E:/data/1-原始数据表/TLT/月度明细/商户维度/'
    choose = input('请输入需汇总数据期间维度(d-日/m-月):')
    if choose == 'd':
        return daily_path
    elif choose == 'm':
        return month_path
    else:
        print('请选择 d/m')
        period_choose()


# 读取所需明细
totalPath = period_choose()
savePath = 'E:/data/2-数据源表/TLT/'
judgeDoc = 'E:/data/2-数据源表/判断条件.xlsx'
dateGet = input('请输入需处理交易明细日期：')
transData = pd.read_excel(totalPath + '{}.xls'.format(dateGet), sheet_name='成功交易统计',
                          usecols=['日期', '商户名称', '商户号', '父商户号', '收入所属方', '一级行业', '二级行业', '交易类型',
                                   '成功笔数(不含跨行)', '成功金额(不含跨行)', '跨行发送银行笔数', '跨行发送银行金额',
                                   '手续费', '成本'])
verifyData = pd.read_excel(totalPath + '{}.xls'.format(dateGet), sheet_name='Sheet1',
                           usecols=['日期', '商户名称', '商户号', '父商户号', '收入所属方', '一级行业', '二级行业', '交易类型',
                                    '交易笔数', '手续费', '成本'])


# 剔除不纳入统计的交易类型
prType = pd.read_excel(judgeDoc, sheet_name='交易类型', index_col='交易类型')
transData = transData[transData['交易类型'].isin(prType.index)]
verifyData = verifyData[verifyData['交易类型'] != '快捷协议签约申请']


# 增加笔数、金额、产品列
transData['笔数'] = transData['成功笔数(不含跨行)'] + transData['跨行发送银行笔数']
transData['金额'] = transData['成功金额(不含跨行)'] + transData['跨行发送银行金额']
transData = pd.merge(transData, prType, how='left', on='交易类型')
verifyData['笔数'] = verifyData['交易笔数']
verifyData['金额'] = 0
verifyData['产品'] = '验证'
transData1 = transData[['日期', '商户名称', '商户号', '父商户号', '收入所属方', '一级行业', '二级行业', '交易类型',
                       '手续费', '成本', '笔数', '金额', '产品']]
verifyData1 = verifyData[['日期', '商户名称', '商户号', '父商户号', '收入所属方', '一级行业', '二级行业', '交易类型',
                         '手续费', '成本', '笔数', '金额', '产品']]

totalData = pd.concat([transData1, verifyData1])


# 卡中心、转账扣款交易剔除重复笔数和金额
require_db = (totalData.loc[:, '二级行业'] == '卡中心') & (totalData.loc[:, '交易类型'] == '转账扣款')
totalData.loc[require_db, '笔数'] = 0
totalData.loc[require_db, '金额'] = 0


# 修改收入所属方
belonging = pd.read_excel(judgeDoc, sheet_name='分润', index_col='商户号')
totalData = pd.merge(totalData, belonging, how='left', on='商户号')
require_bl = (totalData['收入所属方'] == '个人业务事业部') & (totalData['归属分公司'] is not None)
totalData.loc[require_bl, '收入所属方'] = totalData['归属分公司']
totalData.loc[totalData['收入所属方'].isnull(), '收入所属方'] = '直营'
del totalData['归属分公司']


# 增加项目标签、收益
prj = pd.read_excel(judgeDoc, sheet_name='项目', index_col='商户号')
totalData = pd.merge(totalData, prj, how='left', on='商户号')
totalData['收益'] = totalData['手续费'] - totalData['成本']


# 增加商户简称
def cut(x):
    if '卡中心' in x:
        return x[0:x.find('卡中心') + 3]
    elif '分行' in x:
        return x[0:x.find('分行') + 2]
    elif '公司' in x:
        return x[0:x.find('公司') + 2]
    elif '（' in x:
        return x[0:x.find('（')]
    else:
        return x


totalData['商户简称'] = totalData['商户名称'].astype(str).map(cut)


# 调整部分特殊商户简称
totalData.loc[totalData['商户名称'] == '（360借条1）五矿国际信托有限公司', '商户简称'] = '（360借条）五矿国际信托有限公司'
totalData.loc[totalData['商户名称'] == '（360借条2）五矿国际信托有限公司', '商户简称'] = '（360借条）五矿国际信托有限公司'
totalData.loc[totalData['商户名称'] == '中国民生银行股份有限公司信用卡中心', '商户简称'] = '民生银行信用卡中心'
totalData.loc[totalData['商户名称'] == '实时还款', '商户简称'] = '浦东发展银行信用卡中心'


# 商户号类型改为字符串
totalData['商户号'] = totalData['商户号'].astype(str)
totalData['父商户号'] = totalData['父商户号'].astype(str)


# 剔除合计行数据
totalData = totalData[~(totalData['商户名称'].isnull())]


# 汇总数据写入汇总简表
total_dic = {'笔数': totalData['笔数'].sum() / 10000, '金额': totalData['金额'].sum() / 100000000,
             '手续费': totalData['手续费'].sum() / 10000, '收益': totalData['收益'].sum() / 10000}
total_df = pd.DataFrame.from_dict(total_dic, orient='index', columns=['数值'])
total_df.index.name = '指标'
if len(dateGet) == 8:  # 日报表增加累计数据
    lastDate = (datetime.datetime.strptime(dateGet, '%Y%m%d') + datetime.timedelta(days=-1)).strftime('%Y%m%d')
    lastData = pd.read_excel('E:/data/4-日报表&周报表/日报&周报202010/个人业务事业部日报表_{}.xlsx'.format(lastDate),
                             sheet_name='Sheet1', header=1, usecols=['区间', '月累计', '年累计'], nrows=4, index_col='区间')
    if dateGet[-4:] == '0101':
        total_df.loc['笔数', '月累计'] = totalData['笔数'].sum() / 10000
        total_df.loc['金额', '月累计'] = totalData['金额'].sum() / 100000000
        total_df.loc['手续费', '月累计'] = totalData['手续费'].sum() / 10000
        total_df.loc['收益', '月累计'] = totalData['收益'].sum() / 10000
        total_df.loc['笔数', '年累计'] = totalData['笔数'].sum() / 10000
        total_df.loc['金额', '年累计'] = totalData['金额'].sum() / 100000000
        total_df.loc['手续费', '年累计'] = totalData['手续费'].sum() / 10000
        total_df.loc['收益', '年累计'] = totalData['收益'].sum() / 10000
    elif dateGet[-2:] == '01':
        total_df.loc['笔数', '月累计'] = totalData['笔数'].sum() / 10000
        total_df.loc['金额', '月累计'] = totalData['金额'].sum() / 100000000
        total_df.loc['手续费', '月累计'] = totalData['手续费'].sum() / 10000
        total_df.loc['收益', '月累计'] = totalData['收益'].sum() / 10000
        total_df.loc['笔数', '年累计'] = lastData.loc['交易笔数（万）', '年累计'] + totalData['笔数'].sum() / 10000
        total_df.loc['金额', '年累计'] = lastData.loc['交易金额（亿）', '年累计'] + totalData['笔数'].sum() / 100000000
        total_df.loc['手续费', '年累计'] = lastData.loc['手续费（万）', '年累计'] + totalData['笔数'].sum() / 10000
        total_df.loc['收益', '年累计'] = lastData.loc['收益（剔除渠道成本/万）', '年累计'] + totalData['笔数'].sum() / 10000
    else:
        total_df.loc['笔数', '月累计'] = lastData.loc['交易笔数（万）', '月累计'] + totalData['笔数'].sum() / 10000
        total_df.loc['金额', '月累计'] = lastData.loc['交易金额（亿）', '月累计'] + totalData['金额'].sum() / 100000000
        total_df.loc['手续费', '月累计'] = lastData.loc['手续费（万）', '月累计'] + totalData['手续费'].sum() / 10000
        total_df.loc['收益', '月累计'] = lastData.loc['收益（剔除渠道成本/万）', '月累计'] + totalData['收益'].sum() / 10000
        total_df.loc['笔数', '年累计'] = lastData.loc['交易笔数（万）', '年累计'] + totalData['笔数'].sum() / 10000
        total_df.loc['金额', '年累计'] = lastData.loc['交易金额（亿）', '年累计'] + totalData['金额'].sum() / 100000000
        total_df.loc['手续费', '年累计'] = lastData.loc['手续费（万）', '年累计'] + totalData['手续费'].sum() / 10000
        total_df.loc['收益', '年累计'] = lastData.loc['收益（剔除渠道成本/万）', '年累计'] + totalData['收益'].sum() / 10000


# 数据存入电子表格
docSave = pd.ExcelWriter(savePath + 'TLT源表_{}.xlsx'.format(dateGet))
total_df.to_excel(docSave, '汇总数据')
totalData.to_excel(docSave, '汇总明细')
docSave.save()

print('数据处理完成！')
