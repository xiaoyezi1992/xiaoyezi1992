# _*_coding:utf-8_*_
import pandas as pd


# 四行业筛选交易类型,增加笔数、金额、收益、产品、行业、项目标签，并修改联营商户收入所属方（行业分类汇总后手工匹配）
# 二级行业为信用卡中心、交易类型为转账扣款的交易金额、笔数归零；验证剔除快捷协议签约申请交易类型的交易明细
# 业务数据汇总简表，取部分分析用字段明细


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


totalPath = period_choose()
savePath = 'E:/data/2-数据源表/TLT/'
judgeDoc = 'E:/data/2-数据源表/判断条件2.xlsx'
dateGet = input('请输入需处理交易明细日期后缀：')
transData = pd.read_excel(totalPath + dateGet + '.xls', sheet_name='成功交易统计',
                          usecols=['日期', '商户名称', '商户号', '父商户号', '收入所属方', '一级行业', '二级行业','交易类型',
                                   '成功笔数(不含跨行)', '成功金额(不含跨行)', '跨行发送银行笔数', '跨行发送银行金额',
                                   '手续费', '成本'])
verifyData = pd.read_excel(totalPath + dateGet + '.xls', sheet_name='Sheet1',
                           usecols=['日期', '商户名称', '商户号', '父商户号', '收入所属方', '一级行业', '二级行业', '交易类型',
                                    '交易笔数', '手续费', '成本'])


# 交易类型和产品
prType = pd.read_excel(judgeDoc, sheet_name='交易类型', index_col='交易类型')

# 筛选交易类型
transData = transData[transData['交易类型'].isin(prType.index)]
verifyData = verifyData[verifyData['交易类型'] != '快捷协议签约申请']

# 增加笔数、金额、产品列
transData['笔数'] = transData['成功笔数(不含跨行)'] + transData['跨行发送银行笔数']
transData['金额'] = transData['成功金额(不含跨行)'] + transData['跨行发送银行金额']
# for i in transData['交易类型']:
#     transData['产品'] = prType.loc[i, '产品']
verifyData['笔数'] = verifyData['交易笔数']
verifyData['金额'] = 0
verifyData['产品'] = '验证'
# transData1 = transData['日期', '商户名称', '商户号', '父商户号', '收入所属方', '一级行业', '二级行业', '交易类型',
#                        '手续费', '成本', '笔数', '金额', '产品']
# verifyData1 = verifyData['日期', '商户名称', '商户号', '父商户号', '收入所属方', '一级行业', '二级行业','交易类型',
#                          '手续费', '成本', '笔数', '金额', '产品']

totalData = pd.concat([transData1, verifyData1])

# 分润判断、匹配函数
# belonging = pd.read_excel(judgeDoc, sheet_name='分润', index_col='商户号')


# 增加项目标签、收益、行业列，修改收入所属方



# 写入明细


# 汇总数据写入汇总简表



docSave = pd.ExcelWriter(savePath + '试验明细.xlsx')
totalData.to_excel(docSave, '汇总明细')
docSave.save()