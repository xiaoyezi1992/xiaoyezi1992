# coding:utf-8
import xlrd, xlsxwriter

# 将已汇总月度明细按月报表口径筛选
# （两个行业明细汇总并筛选交易类型，成功交易明细中增加笔数金额收益产品行业，并修改计价商户收入所属方）
# 二级行业为信用卡中心、交易类型为转账扣款的交易金额、笔数归零
# 确定数据存储路径并打开源数据文件
dateGet = input('请输入已汇总明细日期后缀（例如202008）：')
totalPath = 'E:/数据/1-综合数据/TLT基础数据/月度明细/商户维度/'
baseData = xlrd.open_workbook(totalPath + ('{}.xls'.format(dateGet)))
pjTable1 = baseData.sheet_by_name('普金')
pjTable2 = baseData.sheet_by_name('普金验证')
yhTable1 = baseData.sheet_by_name('银行')
yhTable2 = baseData.sheet_by_name('银行验证')

# 交易类型判断函数
def tp_judge(val):
    tp_path = 'E:/数据/1-综合数据/TLT基础数据/每日明细/商户维度/判断条件.xlsx'
    tp_data = xlrd.open_workbook(tp_path).sheet_by_name('交易类型')
    tp_list = []
    for i in range(tp_data.nrows):
        tp_list.append(tp_data.cell_value(i,0))
    return val in tp_list

# 筛选数据并汇总明细
dataList1 = []
dataList2 = []
for a in range((pjTable1.nrows - 1)):
    if a == 0:
        dataList1.append(pjTable1.row_values(a))
    else:
        if tp_judge(pjTable1.row_values(a)[pjTable1.row_values(0).index('交易类型')]):
            dataList1.append(pjTable1.row_values(a))
for b in range((pjTable2.nrows - 1)):
    dataList2.append(pjTable2.row_values(b))
for c in range((yhTable1.nrows - 1)):
    if c > 0:
        if tp_judge(yhTable1.row_values(c)[yhTable1.row_values(0).index('交易类型')]):
            dataList1.append(yhTable1.row_values(c))
for d in range((yhTable2.nrows - 1)):
    if d > 0:
        dataList2.append(yhTable2.row_values(d))

# 修改验证明细商户号格式
chg_col = dataList2[0].index('商户号')
for i in range(len(dataList2)):
    if i > 0:
        dataList2[i][chg_col] = str(int(dataList2[i][chg_col]))

# 产品匹配函数
def pr_match(type):
    pr_path = 'E:/数据/1-综合数据/TLT基础数据/每日明细/商户维度/判断条件.xlsx'
    pr_list = xlrd.open_workbook(pr_path).sheet_by_name('交易类型')
    row_num = pr_list.col_values(0).index(type)
    return pr_list.col_values(1)[row_num]

# 行业匹配函数
def ind_match(ind):
    ind_path = 'E:/数据/1-综合数据/TLT基础数据/每日明细/商户维度/判断条件.xlsx'
    ind_list = xlrd.open_workbook(ind_path).sheet_by_name('行业')
    row_num = ind_list.col_values(0).index(ind)
    return ind_list.col_values(1)[row_num]

# 分润判断、匹配函数
def prf_match(No):
    prf_path = 'E:/数据/1-综合数据/TLT基础数据/每日明细/商户维度/判断条件.xlsx'
    prf_list = xlrd.open_workbook(prf_path).sheet_by_name('分润')
    No_list = []
    for row in range(prf_list.nrows):
        if row > 0:
            No_list.append(prf_list.cell_value(row, 0))
    if No in No_list:
        row_num = prf_list.col_values(0).index(No)
        return prf_list.col_values(2)[row_num]

# 增加笔数、金额、收益、产品、行业列，修改收入所属方
num1 = dataList1[0].index('成功笔数(不含跨行)')
num2 = dataList1[0].index('跨行发送银行笔数')
amount1 = dataList1[0].index('成功金额(不含跨行)')
amount2 = dataList1[0].index('跨行发送银行金额')
inc1 = dataList1[0].index('手续费')
cost1 = dataList1[0].index('成本')
inc2 = dataList2[0].index('手续费')
cost2 = dataList2[0].index('成本')
tp_col = dataList1[0].index('交易类型')
ind_col1 = dataList1[0].index('二级行业')
ind_col2 = dataList2[0].index('二级行业')
prf_col1 = dataList1[0].index('商户号')
bl_col1 = dataList1[0].index('收入所属方')
prf_col2 = dataList2[0].index('商户号')
bl_col2 = dataList2[0].index('收入所属方')

for i in range(len(dataList1)):
    if i == 0:
        dataList1[0].append('笔数')
        dataList1[0].append('金额')
        dataList1[0].append('收益')
        dataList1[0].append('产品')
        dataList1[0].append('行业')
    else:
        dataList1[i].append(dataList1[i][num1] + dataList1[i][num2])
        dataList1[i].append(dataList1[i][amount1] + dataList1[i][amount2])
        if dataList1[i][ind_col1] == '信用卡中心' and dataList1[i][tp_col] == '转账扣款':
            dataList1[i][(dataList1[0].index('笔数'))] = 0
            dataList1[i][(dataList1[0].index('金额'))] = 0
        dataList1[i].append(dataList1[i][inc1] - dataList1[i][cost1])
        dataList1[i].append(pr_match(dataList1[i][tp_col]))
        dataList1[i].append(ind_match(dataList1[i][ind_col1]))
        if dataList1[i][bl_col1] == ('普惠金融服务事业部' or '银行服务事业部'):
            if prf_match(dataList1[i][prf_col1]) is not None:
                dataList1[i][bl_col1] = prf_match(dataList1[i][prf_col1])
for i in range(len(dataList2)):
    if i == 0:
        dataList2[0].append('收益')
        dataList2[0].append('产品')
        dataList2[0].append('行业')
    else:
        dataList2[i].append(dataList2[i][inc2] - dataList2[i][cost2])
        dataList2[i].append('验证')
        dataList2[i].append(ind_match(dataList2[i][ind_col2]))
        if dataList2[i][bl_col2] == ('普惠金融服务事业部' or '银行服务事业部'):
            if prf_match(dataList2[i][prf_col2]) is not None:
                dataList2[i][bl_col2] = prf_match(dataList2[i][prf_col2])

# 明细数据写入新工作表
totalExcel = xlsxwriter.Workbook(totalPath + ('月报表明细{}.xlsx'.format(dateGet)))
totalSheet1 = totalExcel.add_worksheet('交易')
totalSheet2 = totalExcel.add_worksheet('验证')
for row, rowData in enumerate(dataList1):
    for col, colData in enumerate(rowData):
        totalSheet1.write(row, col, colData)
for row2, rowData2 in enumerate(dataList2):
    for col2, colData2 in enumerate(rowData2):
        totalSheet2.write(row2, col2, colData2)
totalExcel.close()