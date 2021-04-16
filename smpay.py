# _*_coding:utf-8_*_


import pandas as pd
import os


path = 'E:/data/1-原始数据表/TLT/实名支付/'
period = input('请输入财务收入成本入账月份：')
open_path = path + period + '/'
save_path = 'E:/data/3-结果数据/6-收入&成本核算/' + period + '/'


def total_doc(folder):
    df = pd.DataFrame({})
    for name in os.listdir(folder):
        data = pd.read_excel(folder + '/' + name, header=2, usecols=['客户号', '客户名', '一级行业', '二级行业', '收入所属方',
                                                                     '交易类型', '成功笔数(含跨行)', '成功金额(含跨行)',
                                                                     '计费周期', '应收收入', '已收手续费', '未收手续费'])
        df = pd.concat([df, data])
    return df


detail_org = total_doc(open_path)
detail = detail_org[~(detail_org['客户号'].str.contains('合计：'))]
detail = detail[~(detail['客户号'].str.contains('打印：'))]


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


def str_flt(x):
    if isinstance(x, str):
        x = x.replace(',', '')
    return float(x)


detail['商户简称'] = detail['客户名'].astype(str).map(cut)
# 调整特殊商户简称
detail.loc[detail['商户名称'] == '（360借条1）五矿国际信托有限公司', '商户简称'] = '（360借条）五矿国际信托有限公司'
detail.loc[detail['商户名称'] == '（360借条2）五矿国际信托有限公司', '商户简称'] = '（360借条）五矿国际信托有限公司'
detail.loc[detail['商户名称'] == '中国民生银行股份有限公司信用卡中心', '商户简称'] = '民生银行信用卡中心'
detail.loc[detail['商户名称'] == '实时还款', '商户简称'] = '浦东发展银行信用卡中心'
detail.loc[detail['商户名称'] == '辽宁自贸试验区（营口片区）桔子数字科技有限公司（协议支付）', '商户简称'] = '北京桔子分期电子商务有限公司'  # 20210125更新
detail.loc[detail['商户名称'] == '平安银行股份有限公司信用卡中心1', '商户简称'] = '平安银行信用卡中心'  # 20210401
detail.loc[detail['商户名称'] == '平安银行股份有限公司信用卡中心2', '商户简称'] = '平安银行信用卡中心'  # 20210401
# 增加剔税金额
detail.loc[:, '剔税已收'] = detail['已收手续费'].astype(str).map(str_flt)/1.06
detail.loc[:, '剔税未收'] = detail['未收手续费'].astype(str).map(str_flt)/1.06

detail_gr = detail[detail['一级行业'] == '个人业务']
save_detail = pd.ExcelWriter(save_path + '实名支付明细汇总_{}入账.xlsx'.format(period))
detail.to_excel(save_detail, '全部明细')
detail_gr.to_excel(save_detail, '个人明细')
save_detail.save()
print('--------------------\n' * 5)
