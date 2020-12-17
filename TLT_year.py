# coding:utf-8
import xlrd, xlsxwriter

# 筛选交易类型,验证剔除快捷协议签约申请交易类型的交易明细
# 增加笔数、金额、收益、产品、项目标签，并修改联营商户收入所属方（行业分类汇总后手工匹配）
# 二级行业为卡中心、交易类型为转账扣款的交易金额、笔数归零
# 业务数据汇总简表，取部分分析用字段明细


# 确定数据存储路径并打开源数据文件
dateGet = input('需汇总年度：')
totalPath = 'E:/data/1-原始数据表/TLT/年度明细/'
savePath = 'E:/data/2-数据源表/TLT/'
baseData = xlrd.open_workbook(totalPath + ('{}.xls'.format(dateGet)))
table1 = baseData.sheet_by_name('成功交易统计')
table2 = baseData.sheet_by_name('Sheet1')


# 交易类型判断函数
def tp_judge(val):
    tp_path = 'E:/data/2-数据源表/判断条件.xlsx'
    tp_data = xlrd.open_workbook(tp_path).sheet_by_name('交易类型')
    tp_list = []
    for r in range(tp_data.nrows):
        tp_list.append(tp_data.cell_value(r, 0))
    return val in tp_list


# 筛选交易类型
dataList1 =[]
dataList2 =[]
for a in range((table1.nrows - 1)):
    if a == 0:
        dataList1.append(table1.row_values(a))
    else:
        if tp_judge(table1.row_values(a)[table1.row_values(0).index('交易类型')]):
            dataList1.append(table1.row_values(a))
for b in range((table2.nrows - 1)):
    if table2.row_values(b)[table2.row_values(0).index('交易类型')] == '快捷协议签约申请':
        continue
    else:
        dataList2.append(table2.row_values(b))


# 修改验证明细商户号格式
chg_col = dataList2[0].index('商户号')
chg_col2 = dataList2[0].index('父商户号')
for i in range(len(dataList2)):
    if i > 0:
        dataList2[i][chg_col] = str(int(dataList2[i][chg_col]))
        dataList2[i][chg_col2] = str(int(dataList2[i][chg_col2]))


# 产品匹配函数
def pr_match(tp):
    pr_path = 'E:/data/2-数据源表/判断条件.xlsx'
    pr_list = xlrd.open_workbook(pr_path).sheet_by_name('交易类型')
    row_num = pr_list.col_values(0).index(tp)
    return pr_list.col_values(1)[row_num]


# 项目标签匹配函数
def project_match(num):
    project_path = 'E:/data/2-数据源表/判断条件.xlsx'
    project_list = xlrd.open_workbook(project_path).sheet_by_name('项目')
    if num in project_list.col_values(0):
        row_num = project_list.col_values(0).index(num)
        return project_list.col_values(1)[row_num]


# 分润判断、匹配函数
def prf_match(num):
    prf_path = 'E:/data/2-数据源表/判断条件.xlsx'
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
        dataList1[0].append('项目标签')
    else:
        dataList1[i].append(dataList1[i][dataList1[0].index('成功笔数(不含跨行)')] +
                            dataList1[i][dataList1[0].index('跨行发送银行笔数')])
        dataList1[i].append(dataList1[i][dataList1[0].index('成功金额(不含跨行)')] +
                            dataList1[i][dataList1[0].index('跨行发送银行金额')])
        if dataList1[i][ind_col1] == '卡中心' and dataList1[i][tp_col] == '转账扣款':
            dataList1[i][(dataList1[0].index('笔数'))] = 0
            dataList1[i][(dataList1[0].index('金额'))] = 0
        dataList1[i].append(dataList1[i][dataList1[0].index('手续费')] - dataList1[i][dataList1[0].index('成本')])
        dataList1[i].append(pr_match(dataList1[i][tp_col]))
        if dataList1[i][bl_col1] == ('普惠金融服务事业部' or '银行服务事业部'):
            if prf_match(dataList1[i][prf_col1]) is not None:
                dataList1[i][bl_col1] = prf_match(dataList1[i][prf_col1])
        elif dataList1[i][bl_col1] == '个人业务事业部':
            if prf_match(dataList1[i][prf_col1]) is not None:
                dataList1[i][bl_col1] = prf_match(dataList1[i][prf_col1])
        if dataList1[i][bl_col1] == '普惠金融服务事业部':
            dataList1[i][bl_col1] = '直营'
        elif dataList1[i][bl_col1] == '银行服务事业部':
            dataList1[i][bl_col1] = '直营'
        elif dataList1[i][bl_col1] == '个人服务事业部':
            dataList1[i][bl_col1] = '直营'
        elif dataList1[i][bl_col1] == '个人业务事业部':
            dataList1[i][bl_col1] = '直营'
        dataList1[i].append(project_match(dataList1[i][prf_col1]))

for i in range(len(dataList2)):
    if i == 0:
        dataList2[0].append('收益')
        dataList2[0].append('产品')
        dataList2[0].append('项目标签')
    else:
        dataList2[i].append(dataList2[i][dataList2[0].index('手续费')] - dataList2[i][dataList2[0].index('成本')])
        dataList2[i].append('验证')
        if dataList2[i][bl_col2] == ('普惠金融服务事业部' or '银行服务事业部'):
            if prf_match(dataList2[i][prf_col2]) is not None:
                dataList2[i][bl_col2] = prf_match(dataList2[i][prf_col2])
        elif dataList2[i][bl_col2] == '个人业务事业部':
            if prf_match(dataList2[i][prf_col2]) is not None:
                dataList2[i][bl_col2] = prf_match(dataList2[i][prf_col2])
        if dataList2[i][bl_col2] == '普惠金融服务事业部':
            dataList2[i][bl_col2] = '直营'
        elif dataList2[i][bl_col2] == '银行服务事业部':
            dataList2[i][bl_col2] = '直营'
        elif dataList2[i][bl_col2] == '个人服务事业部':
            dataList2[i][bl_col2] = '直营'
        elif dataList2[i][bl_col2] == '个人业务事业部':
            dataList2[i][bl_col2] = '直营'
        dataList2[i].append(project_match(dataList2[i][prf_col2]))


# 开始写入新工作表
totalExcel = xlsxwriter.Workbook(savePath + ('TLT源表{}.xlsx'.format(dateGet)))
partSheet = totalExcel.add_worksheet('有效字段明细汇总')
# detailSheet1 = totalExcel.add_worksheet('成功交易明细')
# detailSheet2 = totalExcel.add_worksheet('验证明细')

# 写入明细
# for row, rowData in enumerate(dataList1):
#     for col, colData in enumerate(rowData):
#         detailSheet1.write(row, col, colData)
# for row2, rowData2 in enumerate(dataList2):
#     for col2, colData2 in enumerate(rowData2):
#         detailSheet2.write(row2, col2, colData2)


# 所需部分字段明细
field_path = 'E:/data/2-数据源表/判断条件.xlsx'
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