# _*_coding:utf-8_*_
import pandas as pd


# 个人（新）行业筛选交易类型,增加笔数、金额、收益、产品、商户简称，并修改联营商户收入所属方
# 二级行业为卡中心、交易类型为转账扣款的交易金额、笔数归零；验证剔除快捷协议签约申请交易类型的交易明细
# 业务数据汇总简表，汇总分析用字段明细


def period_choose():
    daily_path = 'E:/数据/1-原始数据表/TLT/每日明细/'
    month_path = 'E:/数据/1-原始数据表/TLT/月度明细/商户维度/'
    choose = input('请输入需汇总数据期间维度(d-日/m-月):')
    if choose == 'd':
        return daily_path
    elif choose == 'm':
        return month_path
    else:
        print('请选择 d/m')
        period_choose()


totalPath = period_choose()
savePath = 'E:/数据/2-数据源表/TLT/'
judgeDoc = 'E:/数据/2-数据源表/判断条件2.xlsx'
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


# 分润判断、匹配函数
belonging = pd.read_excel(judgeDoc, sheet_name='分润', index_col='商户号')

#     prf_list = xlrd.open_workbook(prf_path).sheet_by_name('分润')
#     num_list = []
#     for row in range(prf_list.nrows):
#         if row > 0:
#             num_list.append(prf_list.cell_value(row, 0))
#     if num in num_list:
#         row_num = prf_list.col_values(0).index(num)
#         return prf_list.col_values(2)[row_num]
#
#
# 增加笔数、金额、收益、产品列，修改收入所属方



# # 写入明细

#
# # 汇总数据写入汇总简表



# docSave = pd.ExcelWriter(savePath + '试验明细.xlsx')
# transDetail.to_excel(docSave, '成功明细')
# verifyDetail.to_excel(docSave, '验证')
# docSave.save()