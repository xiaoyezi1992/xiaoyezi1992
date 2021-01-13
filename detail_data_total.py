#coding:utf-8
import pandas as pd

# 仅汇总全部成功交易统计明细,增加笔数金额收益列,不做其他处理
# 确定数据存储路径并打开源数据文件
dateGet = input('请输入明细日期后缀（例如202009）：')
detailPath = 'E:/data/1-原始数据表/TLT/月度明细/商户渠道维度/'
savePath = 'E:/data/2-数据源表/TLT/'
pjData = pd.read_excel(detailPath + ('普金{}.xls'.format(dateGet)))
yhData = pd.read_excel(detailPath + ('银行{}.xls'.format(dateGet)))
grData = pd.read_excel(detailPath + ('个人{}.xls'.format(dateGet)))
grNData = pd.read_excel(detailPath + ('个人新{}.xls'.format(dateGet)))


# 汇总明细
totalData = pd.concat([pjData, yhData, grData, grNData])
totalData = totalData[totalData['日期'] != '合计']
totalData['笔数'] = totalData['成功笔数(不含跨行)'] + totalData['跨行发送银行笔数']
totalData['金额'] = totalData['成功金额(不含跨行)'] + totalData['跨行发送银行金额']
totalData['收益'] = totalData['手续费'] - totalData['成本']


# 明细数据写入新工作表
totalExcel = pd.ExcelWriter(savePath + '商户渠道明细汇总测试{}.xls'.format(dateGet))
totalData.to_excel(totalExcel, '汇总明细')
totalExcel.save()
