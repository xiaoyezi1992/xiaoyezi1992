# coding:utf-8
import xlrd, xlsxwriter

# 三行业筛选交易类型,增加笔数、金额、收益、产品，并修改联营商户收入所属方（行业分类汇总后手工匹配）
# 二级行业为信用卡中心、交易类型为转账扣款的交易金额、笔数归零；验证剔除快捷协议签约申请交易类型的交易明细
# 业务数据汇总简表，取部分分析用字段明细

# 确定数据存储路径并打开源数据文件


def time_choose():
    daily_path = 'E:/数据/1-原始数据表/TLT/每日明细/'
    week_path = 'E:/数据/1-原始数据表/TLT/每周明细/'  # 与日报表差异处
    month_path = 'E:/数据/1-原始数据表/TLT/月度明细/商户维度/'  # 与日报表差异处
    choose = input('请输入需汇总数据期间维度(d-日/w-周/m-月):')
    if choose == 'd':
        return daily_path
    elif choose == 'w':
        return week_path
    elif choose == 'm':
        return month_path
    else:
        print('请选择 d/w/m')
        time()


totalPath = time_choose()
savePath = 'E:/数据/2-数据源表/TLT/'
dateGet = input('请输入需汇总明细日期后缀：')
baseData1 = xlrd.open_workbook(totalPath + ('普金{}.xls'.format(dateGet)))
baseData2 = xlrd.open_workbook(totalPath + ('银行{}.xls'.format(dateGet)))
baseData3 = xlrd.open_workbook(totalPath + ('个人{}.xls'.format(dateGet)))

pjTable1 = baseData1.sheet_by_name('成功交易统计')
pjTable2 = baseData1.sheet_by_name('Sheet1')
yhTable1 = baseData2.sheet_by_name('成功交易统计')
yhTable2 = baseData2.sheet_by_name('Sheet1')
grTable1 = baseData3.sheet_by_name('成功交易统计')
grTable2 = baseData3.sheet_by_name('Sheet1')


# 交易类型判断函数
def tp_judge(val):
    tp_path = 'E:/数据/2-数据源表/判断条件.xlsx'
    tp_data = xlrd.open_workbook(tp_path).sheet_by_name('交易类型')
    tp_list = []
    for r in range(tp_data.nrows):
        tp_list.append(tp_data.cell_value(r, 0))
    return val in tp_list


# 筛选数据并汇总明细
dataList1 =[]
dataList2 =[]
for a in range((pjTable1.nrows - 1)):
    if a == 0:
        dataList1.append(pjTable1.row_values(a))
    else:
        if tp_judge(pjTable1.row_values(a)[pjTable1.row_values(0).index('交易类型')]):
            dataList1.append(pjTable1.row_values(a))
for b in range((pjTable2.nrows - 1)):
    if pjTable2.row_values(b)[pjTable2.row_values(0).index('交易类型')] == '快捷协议签约申请':
        continue
    else:
        dataList2.append(pjTable2.row_values(b))
for c in range((yhTable1.nrows - 1)):
    if c > 0:
        if tp_judge(yhTable1.row_values(c)[yhTable1.row_values(0).index('交易类型')]):
            dataList1.append(yhTable1.row_values(c))
for d in range((yhTable2.nrows - 1)):
    if d > 0:
        if yhTable2.row_values(d)[yhTable2.row_values(0).index('交易类型')] == '快捷协议签约申请':
            continue
        else:
            dataList2.append(yhTable2.row_values(d))
for e in range((grTable1.nrows - 1)):
    if e > 0:
        if tp_judge(grTable1.row_values(e)[grTable1.row_values(0).index('交易类型')]):
            dataList1.append(grTable1.row_values(e))
for f in range((grTable2.nrows - 1)):
    if f > 0:
        if grTable2.row_values(f)[grTable2.row_values(0).index('交易类型')] == '快捷协议签约申请':
            continue
        else:
            dataList2.append(grTable2.row_values(f))

# 修改验证明细商户号格式
chg_col = dataList2[0].index('商户号')
chg_col2 = dataList2[0].index('父商户号')
for i in range(len(dataList2)):
    if i > 0:
        dataList2[i][chg_col] = str(int(dataList2[i][chg_col]))
        dataList2[i][chg_col2] = str(int(dataList2[i][chg_col2]))


# 产品匹配函数
def pr_match(type):
    pr_path = 'E:/数据/2-数据源表/判断条件.xlsx'
    pr_list = xlrd.open_workbook(pr_path).sheet_by_name('交易类型')
    row_num = pr_list.col_values(0).index(type)
    return pr_list.col_values(1)[row_num]


# 行业匹配函数
# def ind_match(ind):
#     ind_path = 'E:/数据/2-数据源表/判断条件.xlsx'
#     ind_list = xlrd.open_workbook(ind_path).sheet_by_name('行业')
#     row_num = ind_list.col_values(0).index(ind)
#     return ind_list.col_values(1)[row_num]


# 分润判断、匹配函数
def prf_match(num):
    prf_path = 'E:/数据/2-数据源表/判断条件.xlsx'
    prf_list = xlrd.open_workbook(prf_path).sheet_by_name('分润')
    num_list = []
    for row in range(prf_list.nrows):
        if row > 0:
            num_list.append(prf_list.cell_value(row, 0))
    if num in num_list:
        row_num = prf_list.col_values(0).index(num)
        return prf_list.col_values(2)[row_num]


# 增加笔数、金额、收益、产品列，修改收入所属方
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
        # dataList1[0].append('行业')
    else:
        dataList1[i].append(dataList1[i][dataList1[0].index('成功笔数(不含跨行)')] +
                            dataList1[i][dataList1[0].index('跨行发送银行笔数')])
        dataList1[i].append(dataList1[i][dataList1[0].index('成功金额(不含跨行)')] +
                            dataList1[i][dataList1[0].index('跨行发送银行金额')])
        if dataList1[i][ind_col1] == '信用卡中心' and dataList1[i][tp_col] == '转账扣款':
            dataList1[i][(dataList1[0].index('笔数'))] = 0
            dataList1[i][(dataList1[0].index('金额'))] = 0
        dataList1[i].append(dataList1[i][dataList1[0].index('手续费')] - dataList1[i][dataList1[0].index('成本')])
        dataList1[i].append(pr_match(dataList1[i][tp_col]))
        # dataList1[i].append(ind_match(dataList1[i][ind_col1]))
        if dataList1[i][bl_col1] == ('普惠金融服务事业部' or '银行服务事业部'):
            if prf_match(dataList1[i][prf_col1]) is not None:
                dataList1[i][bl_col1] = prf_match(dataList1[i][prf_col1])

for i in range(len(dataList2)):
    if i == 0:
        dataList2[0].append('收益')
        dataList2[0].append('产品')
        # dataList2[0].append('行业')
    else:
        dataList2[i].append(dataList2[i][dataList2[0].index('手续费')] - dataList2[i][dataList2[0].index('成本')])
        dataList2[i].append('验证')
        # dataList2[i].append(ind_match(dataList2[i][ind_col2]))
        if dataList2[i][bl_col2] == ('普惠金融服务事业部' or '银行服务事业部'):
            if prf_match(dataList2[i][prf_col2]) is not None:
                dataList2[i][bl_col2] = prf_match(dataList2[i][prf_col2])

# 笔数
sum_num = 0
for i in range(len(dataList1)):
    if i > 0:
        sum_num += dataList1[i][dataList1[0].index('笔数')]
for i in range(len(dataList2)):
    if i > 0:
        sum_num += dataList2[i][dataList2[0].index('交易笔数')]
# 金额
sum_amt = 0
for i in range(len(dataList1)):
    if i > 0:
        sum_amt += dataList1[i][dataList1[0].index('金额')]
# 手续费
sum_inc = 0
for i in range(len(dataList1)):
    if i > 0:
        sum_inc += dataList1[i][dataList1[0].index('手续费')]
for i in range(len(dataList2)):
    if i > 0:
        sum_inc += dataList2[i][dataList2[0].index('手续费')]
# 收益
sum_prf = 0
for i in range(len(dataList1)):
    if i > 0:
        sum_prf += dataList1[i][dataList1[0].index('收益')]
for i in range(len(dataList2)):
    if i > 0:
        sum_prf += dataList2[i][dataList2[0].index('收益')]
totalData = ['数据', sum_num/10000, sum_amt/100000000, sum_inc/10000, sum_prf/10000]

# 开始写入新工作表
totalExcel = xlsxwriter.Workbook(savePath + ('TLT数据源表(含个人){}.xlsx'.format(dateGet)))
totalSheet = totalExcel.add_worksheet('业务数据汇总')
partSheet = totalExcel.add_worksheet('有效字段明细汇总')
detailSheet1 = totalExcel.add_worksheet('成功交易明细')
detailSheet2 = totalExcel.add_worksheet('验证明细')

# 写入明细
for row, rowData in enumerate(dataList1):
    for col, colData in enumerate(rowData):
        detailSheet1.write(row, col, colData)
for row2, rowData2 in enumerate(dataList2):
    for col2, colData2 in enumerate(rowData2):
        detailSheet2.write(row2, col2, colData2)

# 汇总数据写入汇总简表
totalRowList = ['项目', '笔数', '金额', '手续费', '收益']
for totalRow, project in enumerate(totalRowList):
    totalSheet.write(totalRow, 0, project)
for totalRow2, data in enumerate(totalData):
    totalSheet.write(totalRow2, 1, data)

# 所需部分字段明细
field_path = 'E:/数据/2-数据源表/判断条件.xlsx'
field_data = xlrd.open_workbook(field_path).sheet_by_name('字段')
field_list1 = []
field_list2 = []
for i in range(field_data.nrows):
    field_list1.append(field_data.cell_value(i, 0))
for i in range(field_data.nrows - 1):
    field_list2.append(field_data.cell_value(i, 2))

for i in range(len(dataList1)):
    for field, j in zip(field_list1,range(len(field_list1))):
        partSheet.write(i, j, dataList1[i][dataList1[0].index(field)])
for i in range(1, len(dataList2)):
    for field2, j in zip(field_list2, range(len(field_list2))):
        partSheet.write(i + len(dataList1) - 1, j, dataList2[i][dataList2[0].index(field2)])
totalExcel.close()