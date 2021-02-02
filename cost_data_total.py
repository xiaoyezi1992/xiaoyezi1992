# coding:utf-8

import os
import pandas as pd


opr_period = input('请输入处理月份（如202101）：')
gr_path = 'E:/data/9-收入&成本核算/{}/银行与普惠金融服务事业部/个人业务部'.format(opr_period)
yp_path = 'E:/data/9-收入&成本核算/{}/银行与普惠金融服务事业部/银行与普惠金融服务事业部'.format(opr_period)
dfs = pd.DataFrame([])
total = pd.DataFrame([])
for gr_name in os.listdir(gr_path):
    gr_total = pd.read_excel(gr_path + '/' + gr_name, index_col=0)
    if gr_name in os.listdir(yp_path):
        yp_total = pd.read_excel(yp_path + '/' + gr_name, index_col=0)
        plus = gr_total + yp_total
        if '汇总表' in gr_name:
            total = pd.concat([total, plus])
        else:
            dfs = dfs.merge(plus, how='outer', left_index=True, right_index=True)
    else:
        dfs = dfs.merge(gr_total, how='outer', left_index=True, right_index=True)
savePath = pd.ExcelWriter('E:/data/9-收入&成本核算/{}/汇总_业管成本{}.xlsx'.format(opr_period, opr_period))
total.to_excel(savePath, '汇总数据')
dfs.to_excel(savePath, '总表汇总')
savePath.save()  # 没有金额
