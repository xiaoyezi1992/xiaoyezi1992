# coding:utf-8

import os
import pandas as pd


opr_period = input('请输入处理月份（如202101）：')
gr_path = 'E:/data/3-结果数据/6-收入&成本核算/{}/银行与普惠金融服务事业部/个人业务部'.format(opr_period)
yp_path = 'E:/data/3-结果数据/6-收入&成本核算/{}/银行与普惠金融服务事业部/银行与普惠金融服务事业部'.format(opr_period)
dfs = pd.DataFrame([])
total = pd.DataFrame([])
company_list = ['银行与普惠金融服务事业部', '个人业务部', '安徽分公司', '北京分公司', '大连分公司', '福建分公司', '甘肃分公司', '广东分公司',
                '汕头分公司', '广西分公司', '贵州分公司', '海南分公司', '河北分公司', '河南分公司', '黑龙江分公司', '湖北分公司', '湖南分公司',
                '吉林分公司', '江苏分公司', '江西分公司', '辽宁分公司', '内蒙古分公司', '宁波分公司', '宁夏分公司', '青岛分公司', '厦门分公司',
                '山东分公司', '山西分公司', '陕西分公司', '上海分公司', '深圳分公司', '四川分公司', '天津分公司', '新疆分公司', '云南分公司',
                '浙江分公司', '重庆分公司', '青海分公司', '西藏分公司']
for gr_name in os.listdir(gr_path):
    gr_total = pd.read_excel(gr_path + '/' + gr_name, index_col=0)
    gr_total.fillna(0)
    if gr_name in os.listdir(yp_path):
        yp_total = pd.read_excel(yp_path + '/' + gr_name, index_col=0)
        yp_total.fillna(0)
        plus = gr_total + yp_total
        if '汇总表' in gr_name:
            total = pd.concat([total, plus])
        else:
            dfs = dfs.merge(plus, how='outer', left_index=True, right_index=True)
    else:
        dfs = dfs.merge(gr_total, how='outer', left_index=True, right_index=True)
savePath = pd.ExcelWriter('E:/data/3-结果数据/6-收入&成本核算/{}/汇总_业管成本{}.xlsx'.format(opr_period, opr_period))
total.to_excel(savePath, '汇总数据')
dfs.to_excel(savePath, '总表汇总')
savePath.save()
