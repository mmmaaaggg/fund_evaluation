#! /usr/bin/env python
# -*- coding:utf-8 -*-
"""
@author  : baoyuan
@Time    : 2018/9/30 13:18
@File    : read_nav_files.py
@contact : mmmaaaggg@163.com
@desc    : 
"""
import os
from fh_utils import try_2_date, date_2_str
import logging
import pandas as pd

logger = logging.getLogger()


def read_nav_files(folder_path_dict: dict):
    fund_dictionay = {}
    folder_path_evaluation_table = folder_path_dict.get('folder_path_evaluation_table')
    file_names = os.listdir(folder_path_evaluation_table)
    for file_name in file_names:
        # file_path = r'd:\Works\F复华投资\合同、协议\丰润\丰润一期\SK8992_复华丰润稳健一期_估值表_20170113.xls'
        file_path = os.path.join(folder_path_evaluation_table, file_name)
        file_name_net, file_extension = os.path.splitext(file_path)
        if file_extension not in ('.xls', '.xlsx'):
            continue
        else:
            logger.debug('file_path: %s', file_path)
            data_df = pd.read_excel(file_path, skiprows=1, header=0).dropna(how='all', axis=0).dropna(how='all', axis=1)
            date = try_2_date(data_df.iloc[0][0])  # data_df.iloc[0][0][-10:]
            # 获取净值
            data_df1 = pd.read_excel(file_path, skiprows=3, header=0)

            # data_df1['科目名称'][data_df1['科目代码'].apply(lambda x: x.find('基金单位净值') != -1 if isinstance(x, str) else False)]
            if '财通' in data_df.columns[0]:
                cum_nav = data_df1['科目名称'][data_df1['科目代码'] == '基金单位净值：']
                name, nav = data_df.columns[0][13:-6], float(cum_nav.values[0])
            elif '万霁' in data_df.columns[0]:
                cum_nav = data_df1['科目名称'][data_df1['科目代码'] == '基金单位净值:']
                name, nav = data_df.columns[0], float(cum_nav.values[0])
            else:
                raise ValueError('请检查文件中的估值表是否发生改变或者里面有“财通”和“万霁”以外的产品')
            fund_dictionay.setdefault(name, []).append([date, nav])

    folder_path_only_nav = folder_path_dict.get('folder_path_only_nav')
    file_names = os.listdir(folder_path_only_nav)
    for file_name in file_names:
        file_path = os.path.join(folder_path_only_nav, file_name)
        file_name_net, file_extension = os.path.splitext(file_path)
        if file_extension not in ('.xls', '.xlsx'):
            continue
        else:
            logger.debug('file_path: %s', file_path)
            data_df = pd.read_excel(file_path, skiprows=0, header=0)
            date = try_2_date(data_df.iloc[0][2])
            name, nav = data_df.iloc[0][0], float(data_df.iloc[0][3])
        # 把新萌的净值加上去
        fund_dictionay.setdefault(name, []).append([date, nav])

    folder_path_cash = folder_path_dict.get('folder_path_cash')
    file_names = os.listdir(folder_path_cash)
    cash_df = None
    for file_name in file_names:
        file_path = os.path.join(folder_path_cash, file_name)
        file_name_net, file_extension = os.path.splitext(file_path)
        if file_extension not in ('.xls', '.xlsx'):
            continue
        else:
            logger.debug('file_path: %s', file_path)
            data_df = pd.read_excel(file_path, skiprows=0, header=0).dropna(how='all')
            name, date, nav = data_df.iloc[0][1:], date_2_str(data_df.iloc[-1][0]), data_df.iloc[-1][1:]
            cash_df = pd.concat([name, nav], axis=1)
            cash_df['date'] = date
            cash_df.columns = ['基金名称', '现金', '日期']
            cash_df.index = range(cash_df.shape[0])

    return fund_dictionay, cash_df


if __name__ == "__main__":
    folder_path = r'd:\WSPych\RefUtils\src\fh_tools\nav_tools\product_nav'
    folder_path_evaluation_table = r'D:\WSPycharm\fund_evaluation\evaluation_table'
    folder_path_only_nav = r'D:\WSPycharm\fund_evaluation\only_nav'
    folder_path_cash = r'D:\WSPycharm\fund_evaluation\cash'
    folder_path_dict = {'folder_path_evaluation_table': folder_path_evaluation_table,
                        'folder_path_only_nav': folder_path_only_nav, 'folder_path_cash': folder_path_cash}
    fund_nav_dic, cash_df = read_nav_files(folder_path_dict)
