# _*_coding:utf-8_*_

import pandas as pd

path = 'E:/data/12-内部计价/'
month = input('请输入月份：')
detail = pd.read_excel(path + month + '/111-内部计价登记表新-{}.xlsx'.format(month), sheet_name=month)
sep_detail = pd.ExcelWriter(path + month + '/内部计价分公司明细{}.xlsx'.format(month))
company = ['安徽分公司', '北京分公司', '大连分公司', '福建分公司', '甘肃分公司', '广东分公司', '广西分公司', '贵州分公司', '海南分公司',
           '河北分公司', '河南分公司', '黑龙江分公司', '湖北分公司', '湖南分公司', '吉林分公司', '江苏分公司', '江西分公司', '辽宁分公司',
           '内蒙古分公司', '宁波分公司', '青岛分公司', '厦门分公司', '山东分公司', '山西分公司', '陕西分公司', '上海分公司', '深圳分公司',
           '四川分公司', '天津分公司', '新疆分公司', '云南分公司', '浙江分公司', '重庆分公司']
for i in company:
    detail_sep1 = detail[detail['转入方'] == i]
    detail_sep2 = detail[detail['转出方'] == i]
    detail_sep = pd.concat([detail_sep1, detail_sep2])
    detail_sep.to_excel(sep_detail, i)
sep_detail.save()