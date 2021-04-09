# coding:utf-8

# 将日报表支付+个人科技明细表按需汇总统计数据
# 源数据：前一日日报表，明细：支付下载统计日、通联钱包下载统计日及当月累计至统计日、到手下载统计日
# 生意金数据为新统计表，其他助贷产品放款及助贷用户下载当天报表，短信引流数据在统一报表平台下载统计日
# daily_table3生意金放款数据调整数据类型,daily_table4增加判断生意金当日是否无数据,daily_table5增加判断当日新增金额是否为0


import pandas as pd
import datetime
import time

start = time.time()
# 确定统计路径和时间
docDate = input('请输入文件下载日期（例如20201030）:')
countDate = input('请输入日报表统计日期:')
tlt_path = 'E:/data/1-原始数据表/TLT/每日明细/'
afterDate = (datetime.datetime.strptime(countDate, '%Y%m%d') + datetime.timedelta(days=1)).strftime('%Y%m%d')
beforeDate = (datetime.datetime.strptime(countDate, '%Y%m%d') + datetime.timedelta(days=-1)).strftime('%Y%m%d')
lastAmtDate = (datetime.datetime.strptime(countDate, '%Y%m%d') + datetime.timedelta(days=-2)).strftime('%Y%m%d')
# 用于生意金已累计放款数据查询
lastMonthDt = (datetime.datetime.strptime(countDate[:6] + '01', '%Y%m%d') + datetime.timedelta(days=-1)).strftime(
    '%Y%m%d')  # 上月末，用于统计助贷用户月累计
lastYearDt = (datetime.datetime.strptime((countDate[:4] + '0101'), '%Y%m%d') + datetime.timedelta(days=-1)).strftime(
    '%Y%m%d')  # 上年末，用于统计助贷用户年累计
dataPath = 'E:/data/1-原始数据表/产品/'
savePath = 'E:/data/2-数据源表/产品/'
lastData = pd.read_excel('E:/data/3-结果数据/1-日报表&周报表/日报&周报202010/个人业务事业部日报表_{}.xlsx'.format(beforeDate),
                         sheet_name='Sheet1', header=1, usecols=['区间', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4',
                                                                 '月累计', '年累计'])
total = pd.ExcelWriter(savePath + '日报表数据{}.xlsx'.format(countDate))


# 支付数据读取
def get_tlt(path, date, last_data):
    trans_data = pd.read_excel(path + '{}.xls'.format(date), sheet_name='成功交易统计',
                               usecols=['日期', '商户名称', '商户号', '父商户号', '收入所属方', '一级行业', '二级行业', '交易类型',
                                        '成功笔数(不含跨行)', '成功金额(不含跨行)', '跨行发送银行笔数', '跨行发送银行金额', '手续费',
                                        '成本'])
    verify_data = pd.read_excel(path + '{}.xls'.format(date), sheet_name='Sheet1',
                                usecols=['日期', '商户名称', '商户号', '父商户号', '收入所属方', '一级行业', '二级行业', '交易类型',
                                         '交易笔数', '手续费', '成本'])
    pr_dict = {'代付': '代付', '实时付款': '代付', '联合付款': '代付', '代收': '代收', '实时收款': '代收', '联合收款': '代收',
               '批量本地身份验证扣款': '代收', '批量身份验证扣款': '代收', '实时本地身份验证扣款': '代收', '实时身份验证扣款': '代收',
               '实时转账': '代收', '快捷协议支付': '快捷', '快捷直接支付': '快捷', '批量协议支付': '快捷', '验证': '验证', '消费': '终端',
               '收银宝': '终端', '卡基': '终端', '扫码': '终端', '信用卡支付': '网关', '网银B2C支付': '网关', '网银B2B支付': '网关',
               '移动APP支付': '网关', 'Wap支付': '网关', '直清还款': '代收', '转账扣款': '代收'}
    trans_data = trans_data[trans_data['交易类型'].isin(pr_dict.keys())]
    verify_data = verify_data[verify_data['交易类型'] != '快捷协议签约申请']
    trans_data['笔数'] = trans_data['成功笔数(不含跨行)'] + trans_data['跨行发送银行笔数']
    trans_data['金额'] = trans_data['成功金额(不含跨行)'] + trans_data['跨行发送银行金额']
    trans_data['产品'] = trans_data['交易类型'].map(pr_dict)
    verify_data['笔数'] = verify_data['交易笔数']
    verify_data['金额'] = 0
    verify_data['产品'] = '验证'
    trans_data1 = trans_data[['日期', '商户名称', '商户号', '父商户号', '收入所属方', '一级行业', '二级行业', '交易类型', '手续费',
                              '成本', '笔数', '金额', '产品']]
    verify_data1 = verify_data[['日期', '商户名称', '商户号', '父商户号', '收入所属方', '一级行业', '二级行业', '交易类型', '手续费',
                                '成本', '笔数', '金额', '产品']]
    total_data = pd.concat([trans_data1, verify_data1])
    require_db = (total_data.loc[:, '二级行业'] == '卡中心') & (total_data.loc[:, '交易类型'] == '转账扣款')
    total_data.loc[require_db, '笔数'] = 0
    total_data.loc[require_db, '金额'] = 0
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
    total_data['归属分公司'] = total_data['商户号'].map(bl_dict)
    require_bl = (total_data['收入所属方'] == '个人业务事业部') & (total_data['归属分公司'].notnull())
    total_data.loc[require_bl, '收入所属方'] = total_data.loc[require_bl, '归属分公司']
    total_data.loc[total_data['收入所属方'] == '个人业务事业部', '收入所属方'] = '直营'
    del total_data['归属分公司']
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
                200440300000081: '360项目', 200440300000021: '360项目', 200440300000004: '360项目', 200440300000002: '360项目',
                200584000025648: '360项目', 200584000025647: '360项目', 200584000025638: '360项目', 200584000025627: '360项目',
                200290000030472: '360项目', 200604000001284: '360项目', 200584000025508: '360项目', 200584000025458: '360项目',
                200584000025369: '360项目', 200584000025192: '360项目', 200584000021685: '360项目', 200290000030176: '美团',
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
                200584000025431: '美团', 200584000025606: '美团', 200584000025607: '美团',
                200440300000003: '美团', 200440300000023: '美团', 200440300000024: '美团', 200440300000025: '美团',
                200584000025435: '美团', 200584000025652: '美团', 200584000025653: '美团'}  # 20210402更新
    total_data['项目标签'] = total_data['商户号'].map(prj_dict)
    total_data['收益'] = total_data['手续费'] - total_data['成本']

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

    total_data['商户简称'] = total_data['商户名称'].astype(str).map(cut)
    total_data.loc[total_data['商户名称'] == '（360借条1）五矿国际信托有限公司', '商户简称'] = '（360借条）五矿国际信托有限公司'
    total_data.loc[total_data['商户名称'] == '（360借条2）五矿国际信托有限公司', '商户简称'] = '（360借条）五矿国际信托有限公司'
    total_data.loc[total_data['商户名称'] == '中国民生银行股份有限公司信用卡中心', '商户简称'] = '民生银行信用卡中心'
    total_data.loc[total_data['商户名称'] == '实时还款', '商户简称'] = '浦东发展银行信用卡中心'
    total_data.loc[total_data['商户名称'] == '辽宁自贸试验区（营口片区）桔子数字科技有限公司（协议支付）', '商户简称'] = '北京桔子分期电子商务有限公司'  # 20210125更新
    total_data.loc[total_data['商户名称'] == '平安银行股份有限公司信用卡中心1', '商户简称'] = '平安银行信用卡中心'  # 20210401
    total_data.loc[total_data['商户名称'] == '平安银行股份有限公司信用卡中心2', '商户简称'] = '平安银行信用卡中心'  # 20210401
    total_data = total_data[total_data['商户名称'].notnull()]

    def deal_str(data):
        data = str(data).split('.')[0]
        return data

    total_data['商户号'] = total_data['商户号'].map(deal_str)
    total_data['父商户号'] = total_data['父商户号'].map(deal_str)
    total_dic = {'笔数': total_data['笔数'].sum() / 10000, '金额': total_data['金额'].sum() / 100000000,
                 '手续费': total_data['手续费'].sum() / 10000, '收益': total_data['收益'].sum() / 10000}
    total_df = pd.DataFrame.from_dict(total_dic, orient='index', columns=[date])
    total_df.index.name = '指标'
    if len(date) == 8:  # 日报表增加累计数据
        if countDate[-4:] == '0101':
            total_df.loc['笔数', '月累计'] = total_data['笔数'].sum() / 10000
            total_df.loc['金额', '月累计'] = total_data['金额'].sum() / 100000000
            total_df.loc['手续费', '月累计'] = total_data['手续费'].sum() / 10000
            total_df.loc['收益', '月累计'] = total_data['收益'].sum() / 10000
            total_df.loc['笔数', '年累计'] = total_data['笔数'].sum() / 10000
            total_df.loc['金额', '年累计'] = total_data['金额'].sum() / 100000000
            total_df.loc['手续费', '年累计'] = total_data['手续费'].sum() / 10000
            total_df.loc['收益', '年累计'] = total_data['收益'].sum() / 10000
        elif countDate[-2:] == '01':
            total_df.loc['笔数', '月累计'] = total_data['笔数'].sum() / 10000
            total_df.loc['金额', '月累计'] = total_data['金额'].sum() / 100000000
            total_df.loc['手续费', '月累计'] = total_data['手续费'].sum() / 10000
            total_df.loc['收益', '月累计'] = total_data['收益'].sum() / 10000
            total_df.loc['笔数', '年累计'] = last_data.iloc[0, 5] + total_data['笔数'].sum() / 10000
            total_df.loc['金额', '年累计'] = last_data.iloc[1, 5] + total_data['金额'].sum() / 100000000
            total_df.loc['手续费', '年累计'] = last_data.iloc[2, 5] + total_data['手续费'].sum() / 10000
            total_df.loc['收益', '年累计'] = last_data.iloc[3, 5] + total_data['收益'].sum() / 10000
        else:
            total_df.loc['笔数', '月累计'] = last_data.iloc[0, 4] + total_data['笔数'].sum() / 10000
            total_df.loc['金额', '月累计'] = last_data.iloc[1, 4] + total_data['金额'].sum() / 100000000
            total_df.loc['手续费', '月累计'] = last_data.iloc[2, 4] + total_data['手续费'].sum() / 10000
            total_df.loc['收益', '月累计'] = last_data.iloc[3, 4] + total_data['收益'].sum() / 10000
            total_df.loc['笔数', '年累计'] = last_data.iloc[0, 5] + total_data['笔数'].sum() / 10000
            total_df.loc['金额', '年累计'] = last_data.iloc[1, 5] + total_data['金额'].sum() / 100000000
            total_df.loc['手续费', '年累计'] = last_data.iloc[2, 5] + total_data['手续费'].sum() / 10000
            total_df.loc['收益', '年累计'] = last_data.iloc[3, 5] + total_data['收益'].sum() / 10000
    return total_df


tlt_data = get_tlt(tlt_path, countDate, lastData)


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


walletUser = get_wallet_user(dataPath, countDate, lastData)


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


loanUser = get_loan_user(dataPath, docDate, countDate, lastMonthDt, lastYearDt, beforeDate)


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


msg_users = get_msg_user(dataPath, countDate, lastData)


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

    # 通联快贷
    tl_data = pd.read_excel((path + '通联快贷贷后数据{}.xlsx'.format(date1)), header=1, usecols=['支用日期', '支用金额'])
    tl_data['支用日期'] = pd.to_datetime(tl_data['支用日期'])
    tl_data.set_index('支用日期', inplace=True)
    tl_data = pd.Series(tl_data['支用金额'], index=tl_data.index)
    tl_amt = tl_data[date2].sum()

    # 生意金
    syj_data = pd.read_excel((path + '生意金汇总数据{}.xlsx'.format(date1)), header=1, usecols=['日期', '当日新增支用 金额', '累计支用金额'])
    syj_data['日期'] = pd.to_datetime(syj_data['日期'], format='%Y%m%d')
    syj_data.set_index('日期', inplace=True)
    # syj_data = pd.Series(syj_data['当日新增支用 金额'], index=syj_data.index)
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

    # 到手商城
    ds_data = pd.read_excel((path + '订单列表{}.xls'.format(date2)), usecols=['订单状态', '订单金额', '期数'])
    ds_data.set_index(['订单状态'], inplace=True)
    list_ds = ['待发货', '已发货', '备货中']
    judge_list = [i in list_ds for i in ds_data.index]
    df_ds = ds_data.loc[judge_list]
    df_ds.set_index('期数', inplace=True)
    ds_amt = int(df_ds.loc[df_ds.index > 0, :].sum())

    # 到手现金借款
    jk_data = pd.read_excel((path + '借款订单列表{}.xls'.format(date2)), usecols=['订单状态', '借款金额'])
    jk_data.set_index(['订单状态'], inplace=True)
    list_jk = ['放款中', '分期还款中', '已完成']
    judge_list_jk = [j in list_jk for j in jk_data.index]
    jk_amt = int(jk_data.loc[judge_list_jk].sum())

    total_amt = (syj_amt + pos_amt + ck_amt + tx_amt + ft_amt + tl_amt + ds_amt + jk_amt) / 10000
    syj_other_amt = (pos_amt + ck_amt + tx_amt + ft_amt + tl_amt) / 10000
    ds_total_amt = (ds_amt + jk_amt) / 10000

    dict_loan_amt = {'新增放款（万）': total_amt,
                     '生意金-网商贷': syj_amt / 10000,
                     '生意金-其他': syj_other_amt,
                     '到手': ds_total_amt}
    df_loan_amt = pd.DataFrame.from_dict(dict_loan_amt, orient='index', columns=[date2])
    df_loan_amt.index.name = '指标'
    if beforeDate[-4:] == '0101':
        df_loan_amt.loc['新增放款（万）', '月累计'] = total_amt
        df_loan_amt.loc['生意金-网商贷', '月累计'] = syj_amt / 10000
        df_loan_amt.loc['生意金-其他', '月累计'] = syj_other_amt
        df_loan_amt.loc['到手', '月累计'] = ds_total_amt
        df_loan_amt.loc['新增放款（万）', '年累计'] = total_amt
        df_loan_amt.loc['生意金-网商贷', '年累计'] = syj_amt / 10000
        df_loan_amt.loc['生意金-其他', '年累计'] = syj_other_amt
        df_loan_amt.loc['到手', '年累计'] = ds_total_amt
    elif beforeDate[-2:] == '01':
        df_loan_amt.loc['新增放款（万）', '月累计'] = total_amt
        df_loan_amt.loc['生意金-网商贷', '月累计'] = syj_amt / 10000
        df_loan_amt.loc['生意金-其他', '月累计'] = syj_other_amt
        df_loan_amt.loc['到手', '月累计'] = ds_total_amt
        df_loan_amt.loc['新增放款（万）', '年累计'] = last.iloc[10, 5] + total_amt
        df_loan_amt.loc['生意金-网商贷', '年累计'] = last.iloc[11, 5] + syj_amt / 10000
        df_loan_amt.loc['生意金-其他', '年累计'] = last.iloc[12, 5] + syj_other_amt
        df_loan_amt.loc['到手', '年累计'] = last.iloc[13, 5] + ds_total_amt
    else:
        df_loan_amt.loc['新增放款（万）', '月累计'] = last.iloc[10, 4] + total_amt
        df_loan_amt.loc['生意金-网商贷', '月累计'] = last.iloc[11, 4] + syj_amt / 10000
        df_loan_amt.loc['生意金-其他', '月累计'] = last.iloc[12, 4] + syj_other_amt
        df_loan_amt.loc['到手', '月累计'] = last.iloc[13, 4] + ds_total_amt
        df_loan_amt.loc['新增放款（万）', '年累计'] = last.iloc[10, 5] + total_amt
        df_loan_amt.loc['生意金-网商贷', '年累计'] = last.iloc[11, 5] + syj_amt / 10000
        df_loan_amt.loc['生意金-其他', '年累计'] = last.iloc[12, 5] + syj_other_amt
        df_loan_amt.loc['到手', '年累计'] = last.iloc[13, 5] + ds_total_amt
    return df_loan_amt


loan_amt = get_loan_amt(dataPath, docDate, beforeDate, lastAmtDate, lastData)
prd = pd.concat([tlt_data, walletUser, loanUser, msg_users, loan_amt])
prd.to_excel(total, '汇总')
total.save()
end = time.time()
print(end - start)
print('-------------------------------------------\n' * 5)