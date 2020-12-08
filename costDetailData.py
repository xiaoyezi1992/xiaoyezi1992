#coding:utf-8
import xlrd, xlsxwriter

# 仅汇总全部成功交易统计明细,增加笔数金额列,不做其他处理
# 确定数据存储路径并打开源数据文件
dateGet = input('请输入四行业明细的日期后缀（例如202009）：')
totalPath = 'E:/data/1-原始数据表/TLT/月度明细/商户渠道维度/'
savePath = 'E:/data/2-数据源表/TLT/'
baseData1 = xlrd.open_workbook(totalPath + ('个人{}.xls'.format(dateGet)))
baseData2 = xlrd.open_workbook(totalPath + ('普金{}.xls'.format(dateGet)))
baseData3 = xlrd.open_workbook(totalPath + ('银行{}.xls'.format(dateGet)))
baseData4 = xlrd.open_workbook(totalPath + ('个人新{}.xls'.format(dateGet)))

grTable = baseData1.sheet_by_name('成功交易统计')
pjTable = baseData2.sheet_by_name('成功交易统计')
yhTable = baseData3.sheet_by_name('成功交易统计')
grNTable = baseData4.sheet_by_name('成功交易统计')


# 汇总明细
dataList =[]
for a in range((grTable.nrows - 1)):
    dataList.append(grTable.row_values(a))
for b in range((pjTable.nrows - 1)):
    if b > 0:
        dataList.append(pjTable.row_values(b))
for c in range((yhTable.nrows - 1)):
    if c > 0:
        dataList.append(yhTable.row_values(c))
for d in range((grNTable.nrows - 1)):
    if d > 0:
        dataList.append(grNTable.row_values(d))

# 增加笔数、金额列
num1 = dataList[0].index('成功笔数(不含跨行)')
num2 = dataList[0].index('跨行发送银行笔数')
amount1 = dataList[0].index('成功金额(不含跨行)')
amount2 = dataList[0].index('跨行发送银行金额')
for i in range(len(dataList)):
    if i == 0:
        dataList[0].append('笔数')
        dataList[0].append('金额')
    else:
        dataList[i].append(dataList[i][num1] + dataList[i][num2])
        dataList[i].append(dataList[i][amount1] + dataList[i][amount2])

# 明细数据写入新工作表
totalExcel = xlsxwriter.Workbook(savePath + ('商户渠道明细{}.xlsx'.format(dateGet)))
totalSheet = totalExcel.add_worksheet('全部明细')
for row, rowData in enumerate(dataList):
    for col, colData in enumerate(rowData):
        totalSheet.write(row, col, colData)
totalExcel.close()