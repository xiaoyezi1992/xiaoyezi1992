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
dateGet = input('请输入需处理交易明细日期：')
transData = pd.read_excel(totalPath + '{}.xls'.format(dateGet), sheet_name='成功交易统计',
                          usecols=['日期', '商户名称', '商户号', '父商户号', '收入所属方', '一级行业', '二级行业', '交易类型',
                                   '成功笔数(不含跨行)', '成功金额(不含跨行)', '跨行发送银行笔数', '跨行发送银行金额',
                                   '手续费', '成本'])
verifyData = pd.read_excel(totalPath + '{}.xls'.format(dateGet), sheet_name='Sheet1',
                           usecols=['日期', '商户名称', '商户号', '父商户号', '收入所属方', '一级行业', '二级行业', '交易类型',
                                    '交易笔数', '手续费', '成本'])


# 剔除不纳入统计的交易类型
pr_dict = {'代付': '代付', '实时付款': '代付', '联合付款': '代付', '代收': '代收', '实时收款': '代收', '联合收款': '代收',
           '批量本地身份验证扣款': '代收', '批量身份验证扣款': '代收', '实时本地身份验证扣款': '代收', '实时身份验证扣款': '代收',
           '实时转账': '代收', '快捷协议支付': '快捷', '快捷直接支付': '快捷', '批量协议支付': '快捷', '验证': '验证', '消费': '终端',
           '收银宝': '终端', '卡基': '终端', '扫码': '终端', '信用卡支付': '网关', '网银B2C支付': '网关', '网银B2B支付': '网关',
           '移动APP支付': '网关', 'Wap支付': '网关', '直清还款': '代收', '转账扣款': '代收'}
transData = transData[transData['交易类型'].isin(pr_dict.keys())]
verifyData = verifyData[verifyData['交易类型'] != '快捷协议签约申请']


# 增加笔数、金额、产品列
transData['笔数'] = transData['成功笔数(不含跨行)'] + transData['跨行发送银行笔数']
transData['金额'] = transData['成功金额(不含跨行)'] + transData['跨行发送银行金额']
transData['产品'] = transData['交易类型'].map(pr_dict)
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
bl_dict = {200584000021614: '深圳分公司', 200584000021615: '深圳分公司', 200584000022727: '深圳分公司', 200584000022726: '深圳分公司',
           200584000021925: '深圳分公司', 200584000022105: '深圳分公司', 200584000024935: '深圳分公司', 200584000025034: '深圳分公司',
           200584000025160: '深圳分公司', 200584000025161: '深圳分公司', 200584000024931: '深圳分公司', 200584000024936: '深圳分公司',
           200584000025014: '深圳分公司', 200584000025035: '深圳分公司', 200584000025153: '深圳分公司', 200584000025210: '深圳分公司',
           200584000024385: '深圳分公司', 200584000023125: '深圳分公司', 200584000023126: '深圳分公司', 200584000024809: '深圳分公司',
           200584000024810: '深圳分公司', 200584000021065: '深圳分公司', 200584000024625: '深圳分公司', 200584000024626: '深圳分公司',
           200690100000581: '深圳分公司', 200604000001024: '深圳分公司', 200584000020705: '深圳分公司', 200584000022725: '深圳分公司',
           200584000022506: '深圳分公司', 200584000021066: '深圳分公司', 200584000021067: '深圳分公司', 200584000021613: '深圳分公司',
           200584000021616: '深圳分公司', 200731000007425: '云南分公司', 200731000007428: '云南分公司', 200731000007424: '云南分公司',
           200731000007427: '云南分公司', 200701000005129: '贵州分公司', 200701000005149: '贵州分公司', 200336000001776: '浙江分公司',
           200336000001796: '浙江分公司', 200665000000161: '四川分公司', 200290000029883: '四川分公司', 200393000011008: '厦门分公司',
           200290000029885: '厦门分公司', 200584000024007: '深圳分公司', 200584000024006: '深圳分公司', 200690100000781: '浙江分公司',
           200193000000901: '内蒙古分公司', 200193000000821: '内蒙古分公司', 200690100000702: '内蒙古分公司',
           200690100000861: '湖北分公司', 200361000009480: '宁波分公司', 200391000008469: '福建分公司', 200611000011925: '江苏分公司',
           200581000019354: '宁波分公司', 200581000018453: '河南分公司', 200602000007539: '宁波分公司', 200100000025547: '安徽分公司',
           200491000020776: '河南分公司', 200491000020777: '河南分公司', 200491000020696: '河南分公司', 200491000020778: '河南分公司',
           200491000020796: '河南分公司', 200290000020807: '大连分公司', 200290000020827: '大连分公司', 200290000020829: '大连分公司',
           200100000020767: '北京分公司', 200290000006143: '上海分公司', 200161000009342: '山西分公司', 200100000007565: '北京分公司',
           200584000004799: '深圳分公司', 200584000022425: '深圳分公司', 200290000004821: '上海分公司', 200452000017993: '宁波分公司',
           200393000008628: '厦门分公司', 200393000011389: '厦门分公司', 200290000027767: '江苏分公司', 200290000029810: '江苏分公司',
           200290000011847: '上海分公司', 200290000011848: '上海分公司', 200290000011849: '上海分公司', 200290000029700: '上海分公司',
           200290000029711: '上海分公司', 200290000029712: '上海分公司', 200290000029487: '上海分公司', 200290000028007: '江苏分公司',
           200290000028567: '河南分公司', 200584000015327: '深圳分公司', 200290000029147: '广西分公司', 200584000020805: '广东分公司',
           200331000016642: '浙江分公司', 200100000004205: '北京分公司', 200100000024828: '北京分公司', 200290000016846: '北京分公司',
           200100000022667: '北京分公司', 200100000019927: '湖北分公司', 200100000023767: '湖北分公司', 200521000009895: '上海分公司',
           200521000009896: '上海分公司', 200584000003526: '深圳分公司', 200192000001968: '内蒙古分公司', 200491000020676: '河南分公司',
           200691900014605: '重庆分公司', 200493000012379: '河南分公司', 200584000023745: '河南分公司', 200651000009825: '四川分公司',
           200581000019154: '河北分公司', 200361000008760: '安徽分公司', 200361000009520: '安徽分公司', 200100000026763: '北京分公司',
           200290000030075: '江苏分公司', 200290000029930: '厦门分公司'}
totalData['归属分公司'] = totalData['商户号'].map(bl_dict)
require_bl = (totalData['收入所属方'] == '个人业务事业部') & (totalData['归属分公司'].notnull())
totalData.loc[require_bl, '收入所属方'] = totalData.loc[require_bl, '归属分公司']
totalData.loc[totalData['收入所属方'] == '个人业务事业部', '收入所属方'] = '直营'
del totalData['归属分公司']


# 增加项目标签、收益
prj_dict = {200584000025456: '360项目', 200584000025437: '360项目', 200584000025434: '360项目', 200584000025375: '360项目',
            200584000025350: '360项目', 200584000025349: '360项目', 200584000025338: '360项目', 200584000025329: '360项目',
            200584000025311: '360项目', 200584000025240: '360项目', 200584000025222: '360项目', 200584000025215: '360项目',
            200584000025191: '360项目', 200584000025190: '360项目', 200584000025189: '360项目', 200584000025172: '360项目',
            200584000025171: '360项目', 200584000025169: '360项目', 200584000025157: '360项目', 200584000025151: '360项目',
            200584000025122: '360项目', 200584000025121: '360项目', 200584000025120: '360项目', 200584000025119: '360项目',
            200584000025118: '360项目', 200584000025117: '360项目', 200584000025116: '360项目', 200584000025115: '360项目',
            200584000025114: '360项目', 200584000025113: '360项目', 200584000025051: '360项目', 200584000025050: '360项目',
            200584000025032: '360项目', 200584000025029: '360项目', 200584000024910: '360项目', 200584000024907: '360项目',
            200584000024826: '360项目', 200584000024820: '360项目', 200584000024806: '360项目', 200584000024767: '360项目',
            200584000024726: '360项目', 200584000024305: '360项目', 200584000024225: '360项目', 200584000024091: '360项目',
            200584000024090: '360项目', 200584000024089: '360项目', 200604000001187: '360项目', 200604000001186: '360项目',
            200604000001185: '360项目', 200604000001184: '360项目', 200584000024088: '360项目', 200584000024087: '360项目',
            200584000024086: '360项目', 200584000024085: '360项目', 200584000024066: '360项目', 200584000023806: '360项目',
            200584000023345: '360项目', 200584000023305: '360项目', 200584000023230: '360项目', 200584000023229: '360项目',
            200584000023228: '360项目', 200584000023227: '360项目', 200584000023225: '360项目', 200584000023145: '360项目',
            200584000023066: '360项目', 200584000022629: '360项目', 200584000022628: '360项目', 200584000022627: '360项目',
            200584000022626: '360项目', 200584000022625: '360项目', 200584000022365: '360项目', 200584000022051: '360项目',
            200584000022050: '360项目', 200584000022049: '360项目', 200584000021945: '360项目', 200584000021549: '360项目',
            200584000021548: '360项目', 200584000021547: '360项目', 200584000021546: '360项目', 200584000021545: '360项目',
            200584000021467: '360项目', 200584000021466: '360项目', 200584000021465: '360项目', 200584000021407: '360项目',
            200584000021406: '360项目', 200584000021405: '360项目', 200584000021309: '360项目', 200584000021308: '360项目',
            200584000021307: '360项目', 200584000021306: '360项目', 200584000021305: '360项目', 200584000020905: '360项目',
            200584000020585: '360项目', 200584000020185: '360项目', 200584000020165: '360项目',  200584000019065: '360项目',
            200584000017945: '360项目', 200316000000381: '360项目', 200584000016566: '360项目', 200290000017932: '360项目',
            200584000024406: '360项目', 200584000023788: '360项目', 200584000021991: '360项目', 200290000029829: '360项目',
            200290000030135: '360项目', 200791000017348: '360项目', 200221000012175: '360项目', 200290000030176: '美团',
            200821000003157: '美团', 200821000003098: '美团', 200821000003097: '美团', 200290000029930: '美团',
            200290000029929: '美团', 200290000029928: '美团', 200821000003079: '美团', 200821000003078: '美团',
            200821000003077: '美团', 200290000029886: '美团', 200290000029885: '美团', 200290000029884: '美团',
            200290000029883: '美团', 200731000007428: '美团', 200731000007427: '美团', 200731000007426: '美团',
            200731000007425: '美团', 200731000007424: '美团', 200731000007423: '美团', 200100000026335: '美团',
            200100000026334: '美团', 200100000026268: '美团', 200100000026267: '美团', 200521000009852: '美团',
            200521000009851: '美团', 200690100000861: '美团', 200690100000841: '美团', 200397000002122: '美团',
            200397000002121: '美团', 200100000026189: '美团', 200100000026188: '美团', 200100000026187: '美团',
            200690100000802: '美团', 200690100000801: '美团', 200138000002725: '美团', 200138000002724: '美团',
            200138000002723: '美团', 200138000002722: '美团', 200690100000781: '美团', 200336000001796: '美团',
            200584000024007: '美团', 200584000024006: '美团', 200584000024005: '美团', 200690100000721: '美团',
            200452000018133: '美团', 200701000005149: '美团', 200701000005129: '美团', 200393000011008: '美团',
            200100000025287: '美团', 200665000000161: '美团', 200100000025207: '美团', 200100000025027: '美团',
            200336000001776: '美团', 200690100000702: '美团', 200690100000701: '美团', 200821000002837: '美团',
            200193000000821: '美团', 200614000002865: '美团', 200305000004192: '美团', 200584000025352: '美团',
            200584000025348: '美团', 200584000025315: '美团', 200584000025247: '美团', 200584000025246: '美团',
            200584000025236: '美团', 200584000025235: '美团', 200584000025220: '美团', 200584000025211: '美团',
            200584000025210: '美团', 200584000025209: '美团', 200584000025194: '美团', 200584000025161: '美团',
            200584000025160: '美团', 200584000025154: '美团', 200584000025153: '美团', 200584000025152: '美团',
            200584000025112: '美团', 200584000025096: '美团', 200584000025039: '美团', 200584000025038: '美团',
            200584000025035: '美团', 200584000025034: '美团', 200584000025014: '美团', 200584000024936: '美团',
            200584000024935: '美团', 200584000024934: '美团', 200584000024933: '美团', 200584000024932: '美团',
            200584000024931: '美团', 200584000024866: '美团', 200584000024822: '美团', 200584000024821: '美团',
            200584000024810: '美团', 200584000024809: '美团', 200584000024627: '美团', 200584000024626: '美团',
            200584000024625: '美团', 200584000024385: '美团', 200584000023985: '美团', 200584000023765: '美团',
            200584000023726: '美团', 200584000023725: '美团', 200584000023126: '美团', 200584000023125: '美团',
            200584000022727: '美团', 200584000022726: '美团', 200584000022725: '美团', 200584000022506: '美团',
            200584000022465: '美团', 200604000001024: '美团', 200584000022106: '美团', 200584000022105: '美团',
            200584000021925: '美团', 200584000021616: '美团', 200584000021615: '美团', 200584000021614: '美团',
            200584000021613: '美团', 200584000021585: '美团', 200584000021067: '美团', 200584000021066: '美团',
            200584000021065: '美团', 200584000020705: '美团', 200691900014265: '美团', 200690100000581: '美团',
            200192000002788: '360项目'}
totalData['项目标签'] = totalData['商户号'].map(prj_dict)
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
totalData.loc[totalData['商户名称'] == '辽宁自贸试验区（营口片区）桔子数字科技有限公司（协议支付）', '商户简称'] = '北京桔子分期电子商务有限公司'  # 20210125更新


# 剔除合计行数据
totalData = totalData[totalData['商户名称'].notnull()]


def deal_str(data):
    data = str(data).split('.')[0]
    return data


totalData['商户号'] = totalData['商户号'].map(deal_str)
totalData['父商户号'] = totalData['父商户号'].map(deal_str)


# 汇总数据写入汇总简表
total_dic = {'笔数': totalData['笔数'].sum() / 10000, '金额': totalData['金额'].sum() / 100000000,
             '手续费': totalData['手续费'].sum() / 10000, '收益': totalData['收益'].sum() / 10000}
total_df = pd.DataFrame.from_dict(total_dic, orient='index', columns=[dateGet])
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

print('完成！完成！完成！\n' * 5)
