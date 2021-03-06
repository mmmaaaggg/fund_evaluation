#! /usr/bin/env python
# -*- coding:utf-8 -*-
"""
@author  : MG
@Time    : 2018/9/17 9:57
@File    : calc_nav.py
@contact : mmmaaaggg@163.com
@desc    : openpyxl, pandas, xlrd
"""
import xlrd
import xlutils.copy
import xlwt
import re
import os
import pandas as pd
from fh_utils import date_2_str, str_2_date, try_2_date, max_id_val, get_first_idx
from datetime import date, datetime
import logging
from read_nav_files import read_nav_files

logger = logging.getLogger()
MIN_DATE = str_2_date('1999-01-01')


def update_nav_file(file_path, fund_nav_dic, cash_dict, nav_date=date.today(), ignore_sheet_set=None):
    """
    更新净值文件中的净值
    :param file_path:
    :param fund_nav_dic:
    :param cash_dict:
    :param nav_date:
    :param ignore_sheet_set:
    :return:
    """
    # nav_date 日期转换
    if nav_date is None:
        nav_date = date.today()
    elif isinstance(nav_date, str):
        nav_date = str_2_date(nav_date)

    logger.info('读取文件：%s', file_path)
    ret_data_list = []
    file_path_name, file_extension = os.path.splitext(file_path)

    workbook = xlrd.open_workbook(file_path)
    sheet_names = workbook.sheet_names()
    available_sheet_name = set()

    # 在源文件基础上增量更新
    file_path_new = file_path_name + '_' + date_2_str(nav_date) + file_extension
    workbook_new = xlutils.copy.copy(workbook)

    # 各个sheet分别处理
    for sheet_num, sheet_name in enumerate(sheet_names, start=0):
        try:
            sheet = workbook.sheet_by_name(sheet_name)
            # 判断第一个cell是不是“基金名称”不是则跳过
            try:
                if sheet.cell_value(0, 0) != '基金名称':
                    continue
            except IndexError:
                logging.warning('sheet [%s] 无效，跳过', sheet_name)
                continue
            if ignore_sheet_set is not None and sheet_name in ignore_sheet_set:
                logger.warning("[%s] 跳过")
                continue
            logger.info('sheet：[%s]', sheet_name)
            available_sheet_name.add(sheet_name)
            # 取得名称，日期，份额数据
            fund_name = sheet.cell_value(0, 1)
            setup_date = xlrd.xldate_as_datetime(sheet.cell_value(1, 1), 0).date()
            fund_volume = sheet.cell_value(2, 1)
            ret_data_dic = {
                'product_name': fund_name,
                'setup_date': setup_date,
                'volume': fund_volume,
                'sub_product_list': [],
            }
            # 读取各种费用及借贷利息，子产品信息
            fee_dic, sub_product_or_loan_info_dic = {}, {}
            row_num = 3
            cell_content = sheet.cell_value(row_num, 0)
            while cell_content != '日期' and cell_content != '':
                type_name = sheet.cell_value(row_num, 0)
                if type_name in ('费用', '费用（按子产品份额）'):
                    # 费用
                    name = sheet.cell_value(row_num, 1)
                    logger.debug('[%s] -> %d:%s', sheet_name, row_num, name)
                    try:
                        base_date = xlrd.xldate_as_datetime(sheet.cell_value(row_num, 5), 0).date()
                    except:
                        base_date = None
                        logger.warning('[%s] -> %d:%s base_date 字段转换失败 %s 无法转换成日期',
                                       sheet_name, row_num, name, str(sheet.cell_value(row_num, 5)))
                    fee_dic[name] = {
                        'name': name,
                        'rate': sheet.cell_value(row_num, 3),
                        'base_date': base_date,
                    }
                    end_date = sheet.cell_value(row_num, 7)
                    if end_date is not None and end_date != '':
                        # 有些管理费，分段计费
                        try:
                            fee_dic[name]['end_date'] = xlrd.xldate_as_datetime(end_date, 0).date()
                        except:
                            logger.warning('[%s] -> %d:%s end_date 字段转换失败 %s 无法转换成日期',
                                           sheet_name, row_num, name, str(end_date))
                    if sheet.cell_value(row_num, 8) == "基准份额":
                        fee_dic[name]['volume_base'] = float(sheet.cell_value(row_num, 9))
                    if sheet.cell_value(row_num, 8) == "固定费用":
                        fee_dic[name]['fix_fee'] = float(sheet.cell_value(row_num, 9))

                elif type_name == '子产品':
                    # 借款，子基金
                    name = sheet.cell_value(row_num, 1)
                    logger.debug('[%s] -> %d:%s', sheet_name, row_num, name)
                    sub_product_or_loan_info_dic[name] = {
                        'name': name,
                        'rate': float(sheet.cell_value(row_num, 3)) if sheet.cell_value(row_num, 3) != '' else None,
                        'base_date': xlrd.xldate_as_datetime(sheet.cell_value(row_num, 5), 0).date() if sheet.cell_value(
                            row_num, 5) != '' else None,
                        'load_cost': float(sheet.cell_value(row_num, 7)) if sheet.cell_value(row_num, 7) != '' else 0,
                        # 部分子产品存在分笔买入的情况，因此要根据产品真实净值计算当前子产品净值
                        'base_prod_name': sheet.cell_value(row_num, 9) if sheet.cell_value(row_num,
                                                                                           8) == '对应产品名称' else None,
                    }
                    # 仅用来输出全部产品名称使用
                    # logger.error(sheet.cell_value(row_num, 9) if sheet.cell_value(row_num, 8) == '对应产品名称' else name)
                else:
                    logger.error('有未识别的行: %d 该行第一列值为：%s', row_num, type_name)

                row_num += 1
                cell_content = sheet.cell_value(row_num, 0)

            # 读取产品名称：横向读取每个产品名称间隔两个cell
            row_start, col_num = row_num, 1
            cell_content = sheet.cell_value(row_num, col_num)
            sub_product_name_list = []
            while cell_content != '':
                sub_product_name_list.append(cell_content)
                col_num += 3
                cell_content = sheet.cell_value(row_num, col_num)
            # 获取历史净值数据
            row_num = row_start + 1
            data_df = pd.read_excel(file_path, sheet_name=sheet_num, header=row_num, index_col=0).reset_index()
            # 计算每个产品最新净值并更新 df 文件
            data_df_new = data_df.append([None]).copy()
            last_row = data_df_new.shape[0] - 1
            data_df_new.iloc[last_row, 0] = nav_date
            tot_val = 0
            for prod_num, sub_product_name in enumerate(sub_product_name_list):
                col_num = 1 + prod_num * 3
                if sub_product_name in sub_product_or_loan_info_dic:
                    sub_product_info_dic = sub_product_or_loan_info_dic[sub_product_name]
                    base_prod_name = sub_product_info_dic['base_prod_name']
                    rate = sub_product_info_dic['rate']
                    if rate is not None:
                        # 借款：计算利息收入加上本金即为市值
                        # 部分特殊产品 rate == 0
                        nav, volume = 1, 0
                        # 市值
                        value = sub_product_info_dic['load_cost'] * (
                                1 +
                                rate * ((nav_date - sub_product_info_dic['base_date']).days / 365)
                        )
                        data_df_new.iloc[last_row, col_num] = nav
                        data_df_new.iloc[last_row, col_num + 1] = volume
                        data_df_new.iloc[last_row, col_num + 2] = value
                        tot_val += value
                    else:
                        # 净值类产品
                        # nav = get_nav(product_name)
                        if base_prod_name is not None and base_prod_name != "" and base_prod_name in fund_nav_dic:
                            # 子基金分批次买入，需要分别找到对应产品的净值，然后计算总市值
                            date_nav_list = fund_nav_dic[base_prod_name]
                            idx, (_, nav) = max_id_val(date_nav_list,
                                                       lambda x: x[0] if x[0] <= nav_date else MIN_DATE)
                            date_latest_new = date_nav_list[idx][0]
                        elif sub_product_name in fund_nav_dic:
                            date_nav_list = fund_nav_dic[sub_product_name]
                            idx, (_, nav) = max_id_val(date_nav_list,
                                                       lambda x: x[0] if x[0] <= nav_date else MIN_DATE)
                            date_latest_new = date_nav_list[idx][0]
                        else:
                            logger.warning("[%s] %s 净值未查到，默认净值为 1", sheet_name,
                                           base_prod_name if base_prod_name is not None and base_prod_name != "" else
                                           sub_product_name)
                            date_latest_new, nav = None, 1

                        # 净值日期匹配检查
                        if date_latest_new is not None and date_latest_new != nav_date:
                            logger.warning("[%s] %s 净值 %.4f 最新净值日期 %s 与 %s 不符，可能存在计算偏差",
                                           sheet_name, sub_product_name, nav, date_latest_new, nav_date)

                        # 净值
                        data_df_new.iloc[last_row, col_num] = nav
                        # 份额不变
                        volume = data_df_new.iloc[last_row - 1, col_num + 1]
                        data_df_new.iloc[last_row, col_num + 1] = volume
                        # 市值
                        if volume == 0:
                            # 如果份额为0，则沿用上一次市值数据
                            value = data_df_new.iloc[last_row - 1, col_num + 2]
                            logger.warning("[%s] %s 份额为0，沿用上一条记录市值数据", sheet_name, sub_product_name)
                        else:
                            value = float(data_df_new.iloc[last_row, col_num + 1]) * nav

                        data_df_new.iloc[last_row, col_num + 2] = value
                        tot_val += value

                    logger.debug('[%s] -> %s [%d] 净值、份额、市值：%.3f%%, %f, %f',
                                 sheet_name, sub_product_name, col_num, nav * 100, volume, value)
                else:
                    nav = data_df_new.iloc[last_row - 1, col_num]
                    volume = data_df_new.iloc[last_row - 1, col_num + 1]
                    value = data_df_new.iloc[last_row - 1, col_num + 2]
                    logger.error('[%s] -> %s 没有相关的基本信息，沿用上一计算日净值、份额、市值：%s',
                                 sheet_name, sub_product_name, (nav, volume, value))
                    data_df_new.iloc[last_row, col_num] = nav
                    data_df_new.iloc[last_row, col_num + 1] = volume
                    data_df_new.iloc[last_row, col_num + 2] = value

                # 保存子产品信息
                ret_data_dic['sub_product_list'].append({
                    'product_name': sub_product_name,
                    'volume': volume,
                    'nav': nav,
                    # 'nav_last': 1.1521,
                    # 'nav_chg': 0.0025,
                    # 'rr': 0.1325,
                    # 'vol_pct': 0.1,  # 持仓比例
                })

            # 更新现金
            if cash_dict is not None and fund_name in cash_dict:
                cash = cash_dict[fund_name]['现金']
                date_latest_new = str_2_date(cash_dict[fund_name]['日期'])
                if date_latest_new != nav_date:
                    logger.warning('[%s] 当前计算净值日期%s 与 %s 产品账户现金统计日期 %s 不符，可能出现计算偏差',
                                   sheet_name, nav_date, fund_name, date_latest_new)
                data_df_new['银行现金'].iloc[last_row] = cash
                tot_val += cash
            else:
                logger.warning('[%s] %s 没有找到现金余额， 使用上一次的数值', sheet_name, fund_name)
                cash = data_df_new['银行现金'].iloc[last_row - 1]
                data_df_new['银行现金'].iloc[last_row] = cash
                tot_val += cash

            # 计算费用
            tot_fee = 0
            for key, info_dic in fee_dic.items():
                if 'fix_fee' in info_dic:
                    manage_fee = info_dic['fix_fee']
                else:
                    end_date = info_dic.setdefault('end_date', nav_date)
                    if end_date > nav_date:
                        end_date = end_date
                    if key not in data_df_new:
                        logger.warning('[%s] %s -> %s 不在费用列表中，可能该笔费用已经结束，将不进行计算 %s', sheet_name, fund_name, key, info_dic)
                        continue
                    volume_base = info_dic['volume_base'] if 'volume_base' in info_dic else fund_volume
                    manage_fee = - (end_date - info_dic['base_date']).days / 365 * volume_base * info_dic['rate']

                data_df_new[key].iloc[last_row] = manage_fee
                tot_fee += manage_fee

            # 计算新净值
            data_df_new['总市值（费前）'].iloc[last_row] = tot_val
            data_df_new['总市值（费后）'].iloc[last_row] = tot_val + tot_fee
            data_df_new['净值（费前）'].iloc[last_row] = tot_val / fund_volume
            data_df_new['净值（费后）'].iloc[last_row] = nav = (tot_val + tot_fee) / fund_volume
            logger.info('[%s] %s 净值（费前）:%.4f 净值（费后）:%.4f',
                        sheet_name, fund_name, tot_val / fund_volume, (tot_val + tot_fee) / fund_volume)

            # 保存文件
            sheet_wt = workbook_new.get_sheet(sheet_num)
            # nav_date
            date_style = xlwt.XFStyle()
            date_style.num_format_str = 'YYYY/M/D'
            sheet_wt.write(row_num + last_row + 1, 0, nav_date, date_style)
            # 各个产品的【净值	份额	市值】
            # 银行现金	管理费1	管理费2	托管费	总市值（费前）	净值（费前）	总市值（费后）	净值（费后）
            # 净值类的数字保留小数点后4位，其他的小数点后2位
            col_name_list = list(data_df_new.columns)
            col_len = data_df_new.shape[1]
            num_style = xlwt.XFStyle()
            num_style.num_format_str = '0.00;[Red]-0.00'
            nav_style = xlwt.XFStyle()
            nav_style.num_format_str = '0.0000;[Red]-0.0000'

            # 更新头部信息表的日期
            row_num_tmp, col_num = 1, 1
            sheet_wt.write(row_num_tmp, col_num, sheet.cell_value(row_num_tmp, col_num), date_style)
            for row_num_tmp in range(3, row_num):
                for col_num in range(4, 8, 2):
                    value = sheet.cell_value(row_num_tmp, col_num)
                    if value is None or (not isinstance(value, str)) or value == "":
                        continue
                    if value.find('日期') == -1 and len(value) <= 6:
                        continue
                    try:
                        sheet_wt.write(row_num_tmp, col_num + 1, sheet.cell_value(row_num_tmp, col_num + 1), date_style)
                    except:
                        pass

            # 设置历史数据格式
            for row_num_tmp in range(row_num + 1, row_num + last_row + 1):
                for col_num in range(0, col_len):
                    try:
                        value = sheet.cell_value(row_num_tmp, col_num)
                        if value is None:
                            continue
                        if col_num == 0:
                            sheet_wt.write(row_num_tmp, col_num, value, date_style)
                        elif col_name_list[col_num].find('净值') >= 0:
                            sheet_wt.write(row_num_tmp, col_num, value, nav_style)
                        else:
                            sheet_wt.write(row_num_tmp, col_num, value, num_style)
                    except:
                        break

            # 设置最新一行数据及格式
            for col_num in range(1, col_len):
                value = data_df_new.iloc[last_row, col_num]
                if value is None:
                    continue
                if col_name_list[col_num].find('净值') >= 0:
                    sheet_wt.write(row_num + last_row + 1, col_num, value, nav_style)
                else:
                    sheet_wt.write(row_num + last_row + 1, col_num, value, num_style)


            # 保存独立 DataFrame 文件
            # file_path_df = file_path_name + '_df_' + date_2_str(nav_date) + file_extension
            # data_df_new.to_excel(file_path_df, sheet_name=sheet_name)
            # 保存返回信息
            ret_data_dic['nav'] = nav
            nav_last_4_parent_product = data_df_new['净值（费后）'].iloc[last_row - 1]
            ret_data_dic['nav_last'] = nav_last_4_parent_product
            ret_data_dic['nav_chg'] = nav - nav_last_4_parent_product
            ret_data_dic['rr'] = nav ** (365 / (nav_date - setup_date).days) if (nav_date - setup_date).days > 10 else 0.0
            ret_data_list.append(ret_data_dic)
        except:
            logger.exception('[%s]净值计算异常', sheet_name)

    # 删除无用 sheet
    all_sheet = workbook_new._Workbook__worksheets
    for sheet_wt in [sheet_wt for sheet_wt in all_sheet if sheet_wt.name not in available_sheet_name]:
        all_sheet.remove(sheet_wt)

    workbook_new._Workbook__worksheets = all_sheet
    # 保持 sheet
    workbook_new.save(file_path_new)

    return ret_data_list, file_path_new


def save_nav_files(data_list, save_path):
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
                style = xlwt.XFStyle()
                style.num_format_str = '_(#,##0_);(#,##0)'
                worksheet.write(row_num, 1, value / 10000, style)
            elif key == "setup_date":
                style = xlwt.XFStyle()
                style.num_format_str = 'YYYY/M/D'
                worksheet.write(row_num, 2, value, style)
            elif key == "nav":
                style = xlwt.XFStyle()
                style.num_format_str = '0.00'
                worksheet.write(row_num, 10, value, style)
            elif key == "nav_last":
                style = xlwt.XFStyle()
                style.num_format_str = '0.00'
                worksheet.write(row_num, 11, value, style)
            elif key == "rr":
                style = xlwt.XFStyle()
                style.num_format_str = '0.00%'
                worksheet.write(row_num, 12, value, style)
            elif key == "nav_chg":
                style = xlwt.XFStyle()
                style.num_format_str = '_(#,##0.00_);[Red](#,##0.00)'
                value = - value
                worksheet.write(row_num, 13, value, style)
            elif key == "sub_product_list":
                for list_item_sub in data_list[0]['sub_product_list']:
                    row_sub_num += 1
                    for key, value in list_item_sub.items():
                        row_real_num = row_num + row_sub_num
                        if key == "product_name":
                            worksheet.write(row_real_num, 3, value)
                        elif key == "volume":
                            style = xlwt.XFStyle()
                            style.num_format_str = '_(#,##0_);(#,##0)'
                            worksheet.write(row_real_num, 4, value / 10000, style)
                        elif key == "nav":
                            style = xlwt.XFStyle()
                            style.num_format_str = '0.00'
                            worksheet.write(row_real_num, 5, value, style)
                        elif key == "nav_last":
                            style = xlwt.XFStyle()
                            style.num_format_str = '0.00'
                            worksheet.write(row_real_num, 6, value, style)
                        elif key == "nav_chg":
                            style = xlwt.XFStyle()
                            style.num_format_str = '_(#,##0.00_);[Red](#,##0.00)'
                            worksheet.write(row_real_num, 7, value, style)
                        elif key == "rr":
                            style = xlwt.XFStyle()
                            style.num_format_str = '0.00%'
                            worksheet.write(row_real_num, 8, value, style)
                        elif key == "vol_pct":
                            style = xlwt.XFStyle()
                            style.num_format_str = '0.00%'
                            worksheet.write(row_real_num, 9, value, style)
            else:
                pass
        if row_sub_num >= 0:
            row_num += row_sub_num
    # 保存
    workbook.save(save_path)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO,
                        format='%(asctime)s %(name)s|%(module)s.%(funcName)s:%(lineno)d %(levelname)s %(message)s')
    # fund_nav_dic, cash_df = None, None
    ignore_sheet_set = {}
    files_folder_path = os.path.join(os.path.abspath(os.path.curdir), 'files')
    folder_path_evaluation_table = os.path.join(files_folder_path, 'evaluation_table')
    folder_path_only_nav = os.path.join(files_folder_path, 'only_nav')
    folder_path_cash = os.path.join(files_folder_path, 'cash')
    folder_path_dict = {'folder_path_evaluation_table': folder_path_evaluation_table,
                        'folder_path_only_nav': folder_path_only_nav,
                        'folder_path_cash': folder_path_cash}
    # fund_nav_dic, cash_dict = None, None
    fund_nav_dic, cash_dict = read_nav_files(folder_path_dict)
    file_path = os.path.join(files_folder_path, '净值 2018-10-10 展弘分红重新计算份额.xls')
    ret_data_list, file_path_new = update_nav_file(file_path, fund_nav_dic, cash_dict, nav_date='2018-9-30',
                                    ignore_sheet_set=ignore_sheet_set)
    save_path = os.path.join(files_folder_path, 'nav_summary.xls')
    save_nav_files(ret_data_list, save_path)
    os.startfile(file_path_new)
