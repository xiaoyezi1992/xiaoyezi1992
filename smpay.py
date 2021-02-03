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
detail.loc[:, '剔税已收'] = detail['已收手续费'].astype(str).map(str_flt)/1.06
detail.loc[:, '剔税未收'] = detail['未收手续费'].astype(str).map(str_flt)/1.06
save_detail = pd.ExcelWriter(save_path + '实名支付明细汇总{}入账.xlsx'.format(period))
detail.to_excel(save_detail, '全部明细')
save_detail.save()
