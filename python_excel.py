import xlwt



# 创建excel工作表
workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('sheet1')
# 设置表头
worksheet.write(0, 0, label='产品名称')  # product_name
worksheet.write(0, 1, label='产品规模（万）')  # volume
worksheet.write(0, 2, label='基金成立日期')  # setup_date
worksheet.write(0, 3, label='所投资管计划/信托计划名称')  # sub_product_list
worksheet.write(0, 4, label='子基金所投规模（万）')  # volume
worksheet.write(0, 5, label='子基金净值')  # nav
worksheet.write(0, 6, label='子基金上期净值')  # nav_last
worksheet.write(0, 7, label='子基金净值变动率')  # nav_chg
worksheet.write(0, 8, label='子基金收益率（年化）')  # rr
worksheet.write(0, 9, label='子基金持仓比例')  # vol_pct
worksheet.write(0, 10, label='基金净值')  # nav
worksheet.write(0, 11, label='上期净值')  # nav_last
worksheet.write(0, 12, label='收益率（年化）')  # rr
worksheet.write(0, 13, label='净值变动率')  # nav_chg
# 将数据写入excel
row_num = 0
for list_item in data_list:
    row_num += 1
    row_sub_num = -1
    for key, value in list_item.items():
        if key == "product_name":
            worksheet.write(row_num, 0, value)

        elif key == "volume":
            worksheet.write(row_num, 1, value)

        elif key == "setup_date":
            worksheet.write(row_num, 2, value)

        elif key == "nav":
            worksheet.write(row_num, 10, value)

        elif key == "nav_last":
            worksheet.write(row_num, 11, value)

        elif key == "rr":
            worksheet.write(row_num, 12, value)

        elif key == "nav_chg":
            worksheet.write(row_num, 13, value)

        elif key == "sub_product_list":
            for list_item_sub in data_list[0]['sub_product_list']:
                row_sub_num += 1
                for key, value in list_item_sub.items():
                    row_real_num = row_num + row_sub_num
                    if key == "product_name":
                        worksheet.write(row_real_num, 3, value)

                    elif key == "volume":
                        worksheet.write(row_real_num, 4, value)

                    elif key == "nav":
                        worksheet.write(row_real_num, 5, value)

                    elif key == "nav_last":
                        worksheet.write(row_real_num, 6, value)

                    elif key == "nav_chg":
                        worksheet.write(row_real_num, 7, value)

                    elif key == "rr":
                        worksheet.write(row_real_num, 8, value)

                    elif key == "vol_pct":
                        worksheet.write(row_real_num, 9, value)

        else:
            pass

    if row_sub_num >= 0:
        row_num += row_sub_num

# 保存
workbook.save(r'D:\OK.xls')

data_list = [
    {
        'product_name': '复华财通定增投资基金',
        'volume': 3924.53,
        'setup_date': '2013/12/31',
        'nav': 1.1492,
        'nav_last': 1.1521,
        'nav_chg': 0.0025,
        'rr': 0.1325,
        'sub_product_list': [
            {
                'product_name': '展弘稳进1号',
                'volume': 400.00,
                'nav': 1.1492,
                'nav_last': 1.1521,
                'nav_chg': 0.0025,
                'rr': 0.1325,
                'vol_pct': 0.1,  # 持仓比例
            },
            {
                'product_name': '新萌亮点1号',
                'volume': 800.00,
                'nav': 1.1592,
                'nav_last': 1.1721,
                'nav_chg': 0.0025,
                'rr': 0.1425,
                'vol_pct': 0.2,  # 持仓比例
            },
        ],
    },
    {
        'product_name': '鑫隆稳进FOF',
        'volume': 3924.53,
        'setup_date': '2013/12/31',
        'sub_product_list': [
            {
                'product_name': '展弘稳进1号',
                'volume': 400.00,
                'nav': 1.1492,
                'nav_last': 1.1521,
                'nav_chg': 0.0025,
                'rr': 0.1325,
                'vol_pct': 0.1,  # 持仓比例
            },
            {
                'product_name': '新萌亮点1号',
                'volume': 800.00,
                'nav': 1.1592,
                'nav_last': 1.1721,
                'nav_chg': 0.0025,
                'rr': 0.1425,
                'vol_pct': 0.2,  # 持仓比例
            },
        ],
        'nav': 1.1492,
        'nav_last': 1.1521,
        'nav_chg': 0.0025,
        'rr': 0.1325,
    },
]
