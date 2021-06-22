# coding:utf-8
import pandas as pd


# 汇总月度商户渠道明细,筛选保留业管核算成本的交易类型，剔除无交易渠道明细，增加代收付类型、笔数、金额列
def detail_sift(name):
    channel_detail = pd.read_excel('E:/data/1-原始数据表/TLT/月度明细/商户渠道维度/{}.xls'.format(name))
    deal_type = {'其他应收款（暂付）': '代付', '代收付手续费支出': '代付', '汇划手续费支出': '代付', '其它费用': '代付', '结算-分润付款': '代付',
                 '网银出金': '代付', '提现': '代付', '实时付款': '代付', '实时转账': '代付', '头寸拨出': '代付', '快速转账': '代付',
                 '标准转账': '代付', '随机验证': '代付', '结算-代收付款': '代付', '结算-代付失败退款': '代付', '结算-代付退票退款': '代付',
                 '结算-代付多余退款': '代付', '代付': '代付', '结算-T+0代收付款': '代付', '联合退款': '代付', '收款分账（付）': '代付',
                 '联合付款': '代付', '付款出资': '代付', '付款归集冲正': '代付', '退票分账出金': '代付', '退款出金': '代付', '退款归集冲正': '代付',
                 '联合收款出金': '代付', '利息收入': '代收', '其他应付款（暂收)': '代收', '其它收入': '代收', '批量本地身份验证扣款': '代收',
                 '终端实时收款': '代收', '批量身份验证扣款': '代收', '网银入金': '代收', '消费': '代收', '转账扣款': '代收', '充值': '代收',
                 '代收': '代收', '实时收款': '代收', '代付扣款': '代收', '商户主动划款记账(补账)': '代收', '实时收款(有磁有密)': '代收',
                 '实时身份验证扣款': '代收', '实时本地身份验证扣款': '代收', '定期定额收款': '代收', '网关支付': '代收', '快捷协议支付': '代收',
                 '快捷直接支付': '代收', '头寸调入': '代收', '收款分账（收）': '代收', '付款归集': '代收', '付款回退': '代收', '联合收款': '代收',
                 '联合退票': '代收', '退票分账入金': '代收', '退款归集': '代收', '退款回退': '代收', '批量协议支付': '代收', '联合收款入金': '代收',
                 '直清还款': '代收'}
    channel_detail['代收付类型'] = channel_detail['交易类型'].map(deal_type)
    channel_detail.dropna(axis=0, subset=['代收付类型'], inplace=True)
    channel_detail.dropna(axis=0, subset=['交易渠道'], inplace=True)
    channel_detail = channel_detail[~((channel_detail['代收付类型'] == '代收') & (channel_detail['交易渠道'] == '手工网银'))]
    channel_detail = channel_detail[~((channel_detail['代收付类型'] == '代付') & (channel_detail['交易渠道'] == '工行'))]
    channel_detail['笔数'] = channel_detail['成功笔数(不含跨行)'] + channel_detail['跨行发送银行笔数']
    channel_detail['金额'] = channel_detail['成功金额(不含跨行)'] + channel_detail['跨行发送银行金额']

    def deal_str(data):
        data = str(data).split('.')[0]
        return data

    channel_detail['商户号'] = channel_detail['商户号'].map(deal_str)
    channel_detail['父商户号'] = channel_detail['父商户号'].map(deal_str)

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

    channel_detail['商户简称'] = channel_detail['商户名称'].astype(str).map(cut)
    # 调整部分特殊商户简称
    channel_detail.loc[channel_detail['商户名称'] == '（360借条1）五矿国际信托有限公司', '商户简称'] = '（360借条）五矿国际信托有限公司'
    channel_detail.loc[channel_detail['商户名称'] == '（360借条2）五矿国际信托有限公司', '商户简称'] = '（360借条）五矿国际信托有限公司'
    channel_detail.loc[channel_detail['商户名称'] == '中国民生银行股份有限公司信用卡中心', '商户简称'] = '民生银行信用卡中心'
    channel_detail.loc[channel_detail['商户名称'] == '实时还款', '商户简称'] = '浦东发展银行信用卡中心'
    channel_detail.loc[channel_detail['商户名称'] == '辽宁自贸试验区（营口片区）桔子数字科技有限公司（协议支付）', '商户简称'] = '北京桔子分期电子商务有限公司'
    # 20210125更新
    channel_detail.loc[channel_detail['商户名称'] == '平安银行股份有限公司信用卡中心1', '商户简称'] = '平安银行信用卡中心'  # 20210401
    channel_detail.loc[channel_detail['商户名称'] == '平安银行股份有限公司信用卡中心2', '商户简称'] = '平安银行信用卡中心'  # 20210401

    return channel_detail


date_get = input('请输入明细日期后缀（例如202009）：')
detail_data = detail_sift(date_get)
totalExcel = pd.ExcelWriter('E:/data/2-数据源表/TLT/商户渠道明细{}_核算成本.xlsx'.format(date_get))
detail_data.to_excel(totalExcel, '全部明细', index=False)
totalExcel.save()
print('---------------\n' * 5)
