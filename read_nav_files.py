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
            for i in range(len(data_df)):
                date = try_2_date(data_df.iloc[i][2])
                name, nav = data_df.iloc[i][0], float(data_df.iloc[i][3])
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
            cash = cash_df.set_index('基金名称')
            cash_0 = cash.drop(['鑫隆稳进FOF-募集户', '鑫隆稳进FOF-托管户', '定增一期 募集户', '定增一期 托管户'])
            cash_1 = pd.DataFrame({'基金名称': ['鑫隆稳进FOF', '复华财通定增投资基金'],
                                   '现金': [cash.loc['鑫隆稳进FOF-募集户']['现金'] + cash.loc['鑫隆稳进FOF-托管户']['现金'],
                                          cash.loc['定增一期 募集户']['现金'] + cash.loc['定增一期 托管户']['现金']]
                                      , '日期': [cash.loc['鑫隆稳进FOF-募集户']['日期'], cash.loc['定增一期 募集户']['日期']]})
            cash = pd.concat([cash_0, cash_1.set_index('基金名称')])
            new_index = ['复华结构化保本混合型基金', '复华结构化保本混合型基金二期', '复华结构化保本混合型基金三期', '复华结构化保本混合型基金四期',
                         '恒银东兴', '长信', '恒银东升', '鑫隆FOF一期','复华财通定增投资基金三期','复华多策略结构化定增基金','鑫隆稳进FOF','复华财通定增投资基金' ]
            cash.index=new_index
            cash_dict = cash.to_dict('index')

            # 鑫隆稳进FOF=float(cash_df[cash_df['基金名称'] == '鑫隆稳进FOF-募集户']['现金']) + \
            # float(cash_df[cash_df['基金名称'] == '鑫隆稳进FOF-托管户']['现金'])
            # list([鑫隆稳进FOF,])
            # 定增一期 = float(cash_df[cash_df['基金名称'] == '定增一期 募集户']['现金']) + \
            #           float(cash_df[cash_df['基金名称'] == '定增一期 托管户']['现金'])
            #
            # #删除掉多余的行
            # y1=cash_df[cash_df['基金名称'].astype(str).str.contains('募集户')]
            # ret=list(set(list(y1['基金名称']))^set(list(cash_df['基金名称'])))
            # cash_df=cash_df[cash_df['基金名称'].isin(ret)]
            # y1 = cash_df[cash_df['基金名称'].astype(str).str.contains( '托管户')]
            # ret = list(set(list(y1['基金名称'])) ^ set(list(cash_df['基金名称'])))
            # cash_df = cash_df[cash_df['基金名称'].isin(ret)]

    return fund_dictionay, cash_dict


if __name__ == "__main__":
    files_folder_path = os.path.join(os.path.abspath(os.path.curdir), 'files')
    folder_path_evaluation_table = os.path.join(files_folder_path, 'evaluation_table')
    folder_path_only_nav = os.path.join(files_folder_path, 'only_nav')
    folder_path_cash = os.path.join(files_folder_path, 'cash')
    folder_path_dict = {'folder_path_evaluation_table': folder_path_evaluation_table,
                        'folder_path_only_nav': folder_path_only_nav,
                        'folder_path_cash': folder_path_cash}
    fund_dictionay, cash_dict = read_nav_files(folder_path_dict)
    print(fund_dictionay)
    print(cash_dict)
