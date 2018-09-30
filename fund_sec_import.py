# -*- coding: utf-8 -*-
"""
Created on 2017/4/5
@author: MG
"""
from datetime import datetime
import pandas as pd
from config_fh import get_db_session, get_db_engine, get_cache_file_path
import xlrd
import re
import os
from enum import IntEnum, unique
from sqlalchemy.types import String, Date

# 判断股票科目
# 1102 股票相关
# 11021199	上交所A股估值增值
# 11023199	深交所A股估值增值
# 11024199	深交所创业板股票估值增值
# 以此开头的所有科目均不统计
RE_PATTERN_STOCK = re.compile(r"(?<=1102\d{2}[0-8]{2})\d{6}")
# re_pattern_ignore = re.compile(r"\d{6}99\w*")  # 11021199123, 31021299IC1703
RE_PATTERN_FUTURE = re.compile(r"(?<=3102\d{2}[0-8]{2})(IC|IF|IH)\d{4}")
RE_PATTERN_FUND = re.compile(r"(?<=1105\d{2}[0-8]{2})\d{6}")
RE_PATTERN_REVERSE_REPO = re.compile(r"(?<=1202\d{4})\d{6}")

@unique
class SecType(IntEnum):
    STOCK = 0
    FUTURE = 1
    BOND = 2
    REVERSE_REPO = 3
    FUND = 4


@unique
class OrderDirection(IntEnum):
    LONG = 1
    SHORT = 0


@unique
class SourceMark(IntEnum):
    WIND = 0
    MANUAL = 1
    IMPORT = 2


def get_fund_nav_date(file_path):
    data_xls = xlrd.open_workbook(file_path)  # 打开xls文件
    # sheet_name = data.sheet_names()[0]
    data_sheet = data_xls.sheet_by_index(0)
    val_cell = data_sheet.cell(2, 0).value
    if val_cell.find('估值日期：') != 0:
        raise('估值日期未找到：%s' % (file_path))
    nav_date = datetime.strptime(val_cell[5:], '%Y年%m月%d日').date()
    return nav_date


def filter_trade_status_str(status_str):
    if type(status_str) != str:
        return ''
    elif status_str in ('【正常交易】', '【按成本估值】'):
        return ''
    else:
        m = re.search(r"(?<=【停牌】\()\d{4}-\d{1,2}-\d{1,2}", status_str)
        if m is not None:
            return m.group()
        else:
            return status_str


def filter_trade_status(hold_pos_df):
    trade_status_s = hold_pos_df['trade_status']
    result = trade_status_s.apply(filter_trade_status_str)
    hold_pos_df['trade_status'] = result
    return hold_pos_df

NAME_CODE_DIC = None
TRADE_CODE_WIND_CODE_DIC = None

def get_stock_name_code_dic():
    global NAME_CODE_DIC
    if NAME_CODE_DIC is None:
        with get_db_session() as session:
            table = session.execute('select sec_name, wind_code from wind_stock_info')
            NAME_CODE_DIC = dict(table.fetchall())
        # print('name_code_dic:\n', name_code_dic)
    return NAME_CODE_DIC


def get_trade_code_wind_code_dic():
    global TRADE_CODE_WIND_CODE_DIC
    if TRADE_CODE_WIND_CODE_DIC is None:
        with get_db_session() as session:
            table = session.execute('select trade_code, wind_code from wind_stock_info')
            TRADE_CODE_WIND_CODE_DIC = dict(table.fetchall())
        # print('name_code_dic:\n', name_code_dic)
    return TRADE_CODE_WIND_CODE_DIC


def get_stock_holding_df(df):

    title_nav_str = '今日单位净值:'
    title_nav_acc_str = '累计单位净值:'
    title_nav_tot = '今日资产净值:'

    name_code_dic = get_stock_name_code_dic()
    df_count = df.shape[0]
    subject_name_s = df['科目名称']
    subject_code_s = df['科目代码']
    price_value_pct_s = df['市值占净值%']
    # 11021199	上交所A股估值增值
    # 11023199	深交所A股估值增值
    # 11024199	深交所创业板股票估值增值
    # 以此开头的所有科目均不统计
    re_pattern_ignore = re.compile(r"\d{6}99\w*")  # 11021199123, 31021299IC1703
    row_num_list = []
    stock_code_list = []
    ret_dic = {}
    for n in range(df_count):
        subject_code = subject_code_s[n]
        # 科目代码为空不记录
        if type(subject_code) != str:
            continue
        if subject_code == '110241':
            ret_dic['cyb_pct'] = price_value_pct_s[n]
        # 科目代码为 11021199 等开头的均不记录 11021199123, 31021299IC1703
        if re_pattern_ignore.match(subject_code) is not None:
            continue
        stock_name = subject_name_s[n]
        if stock_name in name_code_dic:
            row_num_list.append(n)
            stock_code_list.append(name_code_dic[stock_name])
        # 获取今日净值
        if subject_code.find(title_nav_str) == 0:
            ret_dic['nav'] = subject_name_s[n]
        # 获取今日累计净值
        if subject_code.find(title_nav_acc_str) == 0:
            ret_dic['nav_acc'] = subject_name_s[n]
        # 获取今日资产净值
        if subject_code.find(title_nav_tot) == 0:
            ret_dic['nav_tot'] = subject_name_s[n]

    # 修改股票名称为股票代码
    if len(row_num_list) == 0:
        print('stock is []')
    else:
        df.loc[row_num_list, '科目名称'] = stock_code_list

    # 获取完整持仓列表
    hold_pos_df = df.loc[row_num_list, :]
    # 增加 sec_type 列信息
    hold_pos_df['sec_type'] = SecType.STOCK
    # 增加 direction 列信息
    hold_pos_df['direction'] = OrderDirection.LONG

    rename_dic = {'科目名称': 'sec_code',
                  '数量': 'position',
                  '单位成本': 'cost_unit',
                  '成本': 'cost_tot',
                  '成本占净值%': 'cost_pct',
                  '市值': 'value_tot',
                  '市值占净值%': 'value_pct',
                  '停牌信息': 'trade_status',
                  }
    hold_pos_df = hold_pos_df.rename(columns=rename_dic)
    hold_pos_df = hold_pos_df[['sec_code',
                               'direction',
                               'position',
                               'cost_unit',
                               'cost_tot',
                               'cost_pct',
                               'value_tot',
                               'value_pct',
                               'trade_status',
                               'sec_type']]
    hold_pos_df = filter_trade_status(hold_pos_df)
    is_duplicate = hold_pos_df.duplicated('sec_code')
    if any(is_duplicate):
        hold_pos_df = hold_pos_df.loc[not is_duplicate]
    return hold_pos_df, ret_dic


def get_index_holding_df(df, net_short_holding=True):
    df_count = df.shape[0]
    # 股指多单 科目代码头信息
    header_subject_str = '31021101'

    # subject_name_s = df['科目名称']
    subject_code_s = df['科目代码']
    cost_unit_s = df['单位成本']
    re_pattern_index = re.compile(r"(?<=3102\d{4})(IC1|IF1|IH1)[0-9]{3}")  # 3102 + 四位数字 + IC1/IH1/IF1 + 三位数字
    # 11021199	上交所A股估值增值
    # 11023199	深交所A股估值增值
    # 11024199	深交所创业板股票估值增值
    # 以此开头的所有科目均不统计
    re_pattern_ignore = re.compile(r"\d{6}99\w*")  # 11021199123, 31021299IC1703
    row_num_list = []
    df_index_name = []
    direction_list = []
    contract_row_num_dic = {}
    for n in range(df_count):
        subject_code = subject_code_s[n]
        # 科目代码为空不记录
        if type(subject_code) != str:
            continue
        # 科目代码为 11021199 等开头的均不记录 11021199123, 31021299IC1703
        if re_pattern_ignore.match(subject_code) is not None:
            continue
        # 获取股指期货信息
        m = re_pattern_index.search(subject_code)
        if m is not None:
            # print('subject_code', subject_code)
            row_num_list.append(n)
            contract_code = m.group()
            df_index_name.append(contract_code)
            if contract_code in contract_row_num_dic:
                row_num_pair = contract_row_num_dic[contract_code]
            else:
                row_num_pair = [-1, -1]
            if cost_unit_s[n] >= 0:
                direction_list.append(OrderDirection.LONG)
                row_num_pair[0] = n
            else:
                direction_list.append(OrderDirection.SHORT)
                row_num_pair[1] = n
            contract_row_num_dic[contract_code] = row_num_pair

    # 默认持仓方向
    df['direction'] = OrderDirection.SHORT

    row_num_count = len(row_num_list)
    # 增加股指期货名称
    if row_num_count == 0:
        print('index is []')
    else:
        df.loc[row_num_list, '科目名称'] = df_index_name
        df.loc[row_num_list, 'sec_type'] = SecType.FUTURE
        df.loc[row_num_list, 'direction'] = direction_list

    # 获取完整持仓列表
    hold_pos_df = df.loc[row_num_list, :]
    # 增加 sec_type 列信息
    hold_pos_df['sec_type'] = SecType.FUTURE

    if net_short_holding and row_num_count > 0:
        net_holding = hold_pos_df.drop_duplicates('科目名称', keep=False)
        for contract_code, row_num_pair in contract_row_num_dic.items():
            row_num_long, row_num_short = row_num_pair
            if row_num_long < 0 or row_num_short < 0:
                continue
            long_s = df.iloc[row_num_long, :]
            short_s = df.iloc[row_num_short, :].copy()
            short_s['数量'] -= long_s['数量']
            short_s[['成本', '成本占净值%', '市值', '市值占净值%']] += long_s[['成本', '成本占净值%', '市值', '市值占净值%']]
            net_holding = net_holding.append(short_s)
        hold_pos_df = net_holding

    rename_dic = {'科目名称': 'sec_code',
                  '数量': 'position',
                  '单位成本': 'cost_unit',
                  '成本': 'cost_tot',
                  '成本占净值%': 'cost_pct',
                  '市值': 'value_tot',
                  '市值占净值%': 'value_pct',
                  '停牌信息': 'trade_status',
                  }
    hold_pos_df = hold_pos_df.rename(columns=rename_dic)
    hold_pos_df = hold_pos_df[['sec_code',
                               'direction',
                               'position',
                               'cost_unit',
                               'cost_tot',
                               'cost_pct',
                               'value_tot',
                               'value_pct',
                               'trade_status',
                               'sec_type']]
    hold_pos_df = filter_trade_status(hold_pos_df)
    # is_duplicate = hold_pos_df.duplicated('sec_code')
    # if any(is_duplicate):
    #     hold_pos_df.to_csv(file_path + '.csv')
    #     hold_pos_df = hold_pos_df.loc[not is_duplicate]
    return hold_pos_df


def get_holding_df(file_path):
    # file_path = r'd:\Works\F复华投资\合同、协议\丰润\丰润一期\SK8992_复华丰润稳健一期_估值表_20170113.xls'
    # 获取股票资产列表
    df = pd.read_excel(file_path, skiprows=3, header=0)

    stock_holding_df, ret_dic = get_stock_holding_df(df)

    index_holding_df = get_index_holding_df(df)


    return pd.concat([stock_holding_df, index_holding_df]), ret_dic


def get_stock_data(data_s):
    global RE_PATTERN_STOCK
    subject_code = data_s['科目代码']
    m = RE_PATTERN_STOCK.search(subject_code)
    trade_code_wind_code_dic = get_trade_code_wind_code_dic()
    wind_code = trade_code_wind_code_dic[m.group()]
    ret_dic = {'sec_code': wind_code,
               'direction': OrderDirection.LONG,
               'position': data_s['数量'],
               'cost_unit': data_s['单位成本'],
               'cost_tot': data_s['成本'],
               'cost_pct': data_s['成本占净值%'],
               'value_tot': data_s['市值'],
               'value_pct': data_s['市值占净值%'],
               'trade_status': data_s['停牌信息'],
               'sec_type': SecType.STOCK,
    }
    return ret_dic


def get_future_data(data_s):
    subject_code = data_s['科目代码']
    m = RE_PATTERN_FUTURE.search(subject_code)
    cost_unit = data_s['单位成本']
    ret_dic = {'sec_code': m.group(),
               'direction': OrderDirection.LONG if cost_unit >= 0 else OrderDirection.SHORT,
               'position': data_s['数量'],
               'cost_unit': cost_unit,
               'cost_tot': data_s['成本'],
               'cost_pct': data_s['成本占净值%'],
               'value_tot': data_s['市值'],
               'value_pct': data_s['市值占净值%'],
               'trade_status': data_s['停牌信息'],
               'sec_type': SecType.FUTURE,
    }
    return ret_dic


def get_reverse_repo_data(data_s):
    subject_code = data_s['科目代码']
    m = RE_PATTERN_REVERSE_REPO.search(subject_code)
    ret_dic = {'sec_code': m.group(),
               'direction': OrderDirection.LONG,
               'position': data_s['数量'],
               'cost_unit': data_s['单位成本'],
               'cost_tot': data_s['成本'],
               'cost_pct': data_s['成本占净值%'],
               'value_tot': data_s['市值'],
               'value_pct': data_s['市值占净值%'],
               'trade_status': data_s['停牌信息'],
               'sec_type': SecType.REVERSE_REPO,
    }
    return ret_dic


def get_fund_data(data_s):
    subject_code = data_s['科目代码']
    m = RE_PATTERN_FUND.search(subject_code)
    ret_dic = {'sec_code': m.group(),
               'direction': OrderDirection.LONG,
               'position': data_s['数量'],
               'cost_unit': data_s['单位成本'],
               'cost_tot': data_s['成本'],
               'cost_pct': data_s['成本占净值%'],
               'value_tot': data_s['市值'],
               'value_pct': data_s['市值占净值%'],
               'trade_status': data_s['停牌信息'],
               'sec_type': SecType.FUND,
    }
    return ret_dic


def get_sec_data_list(file_path):
    # file_path = r'd:\Works\F复华投资\合同、协议\丰润\丰润一期\SK8992_复华丰润稳健一期_估值表_20170113.xls'

    global RE_PATTERN_STOCK
    global RE_PATTERN_FUTURE
    global RE_PATTERN_FUND
    global RE_PATTERN_REVERSE_REPO

    # 获取股票资产列表
    data_df = pd.read_excel(file_path, skiprows=3, header=0)
    # stock_holding_df, ret_dic = get_stock_holding_df(df)
    # index_holding_df = get_index_holding_df(df)
    data_count = data_df.shape[0]
    sec_data_list = []

    title_nav_str = '今日单位净值:'
    title_nav_acc_str = '累计单位净值:'
    title_nav_tot = '今日资产净值:'
    ret_dic = {}

    for n in range(data_count):
        data_s = data_df.iloc[n, :]
        # 检查 科目代码、科目名称，并绝对以哪一个函数进行数据处理
        subject_code = data_s['科目代码']
        subject_name = data_s['科目名称']
        if type(subject_code) is str:
            # 股票
            m = RE_PATTERN_STOCK.search(subject_code)
            if m is not None:
                # stock_code = m.group()
                sec_data = get_stock_data(data_s)
                sec_data_list.append(sec_data)
                continue
            # 股指期货
            m = RE_PATTERN_FUTURE.search(subject_code)
            if m is not None:
                # stock_code = m.group()
                sec_data = get_future_data(data_s)
                sec_data_list.append(sec_data)
                continue
            # 逆回购
            m = RE_PATTERN_REVERSE_REPO.search(subject_code)
            if m is not None:
                # stock_code = m.group()
                sec_data = get_reverse_repo_data(data_s)
                sec_data_list.append(sec_data)
                continue
            # 基金
            m = RE_PATTERN_FUND.search(subject_code)
            if m is not None:
                # stock_code = m.group()
                sec_data = get_fund_data(data_s)
                sec_data_list.append(sec_data)
                continue
            # 净值数据
            # 获取今日净值
            if subject_code.find(title_nav_str) == 0:
                ret_dic['nav'] = subject_name
                continue
            # 获取今日累计净值
            if subject_code.find(title_nav_acc_str) == 0:
                ret_dic['nav_acc'] = subject_name
                continue
            # 获取今日资产净值
            if subject_code.find(title_nav_tot) == 0:
                ret_dic['nav_tot'] = subject_name
                continue
            # 现金
            # 暂时不做
        else:
            # 科目表为其他字段，暂时忽略
            pass

    return sec_data_list, ret_dic


def import_fund_sec_pct(wind_code, folder_path, mode='delete_insert'):
    file_names = os.listdir(folder_path)
    # file_names = [os.path.join(folder_path, 'SK8992_复华丰润稳健一期_估值表_20170124.xls')]
    nav_date_list, nav_list, nav_acc_list, nav_tot_list, hold_pos_df_list = [], [], [], [], []
    cyb_pct_list = []
    for file_name in file_names:
        # file_path = r'd:\Works\F复华投资\合同、协议\丰润\丰润一期\SK8992_复华丰润稳健一期_估值表_20170113.xls'
        file_path = os.path.join(folder_path, file_name)
        file_name_net, file_extension = os.path.splitext(file_path)
        if file_extension != '.xls':
            print('ignore:', file_name)
            continue
        else:
            print('handle:', file_name)

        # 获取净值日期
        nav_date = get_fund_nav_date(file_path)
        # 获取股票持仓
        # hold_pos_df, ret_dic = get_holding_df(file_path)
        sec_data_list, ret_dic = get_sec_data_list(file_path)
        # 增加日期列
        for dic in sec_data_list:
            dic['nav_date'] = nav_date
        # hold_pos_df['nav_date'] = nav_date

        # 增加净值数据
        # print(nav_date, nav, nav_acc, hold_pos_df)
        nav_date_list.append(nav_date)
        nav_list.append(ret_dic['nav'])
        nav_acc_list.append(ret_dic['nav_acc'])
        nav_tot_list.append(ret_dic['nav_tot'])
        hold_pos_df_list.extend(sec_data_list)
        # 临时提取创业板指仓位占比
        # cyb_pct_list.append(ret_dic.setdefault('cyb_pct', 0))
    for dic in hold_pos_df_list:
        dic['trade_status'] = filter_trade_status_str(dic['trade_status'])
    # 临时提取创业板指仓位占比
    # cyb_pct_df = pd.DataFrame({'nav_date': nav_date_list,
    #                            'cyb_pct': cyb_pct_list,
    #                            })
    # cyb_pct_df.set_index('nav_date', inplace=True)
    # cyb_pct_df.to_csv('cyb_pct_df.csv')
    if mode is None:
        return

    engine = get_db_engine()
    # 更新净值
    nav_df = pd.DataFrame({'nav_date': nav_date_list,
                           'nav': nav_list,
                           'nav_acc': nav_acc_list,
                           'nav_tot': nav_tot_list})
    nav_df['wind_code'] = wind_code
    nav_df['source_mark'] = SourceMark.IMPORT
    nav_df.set_index(['nav_date', 'wind_code'], inplace=True)
    # nav_df.to_csv('fr_nav.csv')
    if mode == 'delete_insert':
        sql_str = 'delete from fund_nav where wind_code=:wind_code'
        with get_db_session(engine) as session:
            session.execute(sql_str, {"wind_code": wind_code})
            print('clean fund_nav data on %s' % wind_code)
    nav_df.to_sql('fund_nav', engine, if_exists='append')
    print('%d data was imported into fund_nav on %s' % (nav_df.shape[0], wind_code))

    # hold_pos_df_tot = pd.concat(hold_pos_df_list)
    # hold_pos_df_tot['wind_code'] = wind_code
    # hold_pos_df_tot.set_index(['wind_code', 'sec_code', 'nav_date', 'direction'], inplace=True)
    # print(hold_pos_df_tot.shape)
    # print(hold_pos_df_tot.index)
    # hold_pos_df_tot.to_csv('asdf.csv')
    # 删除历史数据
    if mode == 'delete_insert':
        sql_str = 'delete from fund_sec_pct where wind_code=:wind_code'
        with get_db_session(engine) as session:
            session.execute(sql_str, {"wind_code": wind_code})
            print('clean fund_sec_pct data on %s' % wind_code)
    # hold_pos_df_tot.to_sql('fund_sec_pct', engine, if_exists='append')
    # print('%d data imported into fund_sec_pct on %s', hold_pos_df_tot.shape[0], wind_code)
    # stock_df = get_stock_df(file_path)
    # print(stock_df)
    # 导入估值表
    for dic in hold_pos_df_list:
        dic["wind_code"] = wind_code
    sec_df = pd.DataFrame(hold_pos_df_list).set_index(['wind_code', 'sec_code', 'nav_date', 'direction'])
    sec_df.to_sql('fund_sec_pct', engine, if_exists='append')
    #     with get_db_session(engine) as session:
    #         sql_str = """INSERT INTO fund_sec_pct (wind_code, sec_code, nav_date, direction, position, cost_unit, cost_tot, cost_pct, value_tot, value_pct, trade_status, sec_type)
    # VALUES (:wind_code, :sec_code, :nav_date, :direction, :position, :cost_unit, :cost_tot, :cost_pct, :value_tot, :value_pct, :trade_status, :sec_type)"""
    #         session.execute(sql_str, hold_pos_df_list)
    print('%d data was imported into fund_sec_pct on %s' % (len(hold_pos_df_list), wind_code))


def export_nav_excel(wind_code):
    """
    导出指定基金的净值数据到excel文件
    :param wind_code: 
    :return: 
    """
    engine = get_db_engine()
    sql_str = 'SELECT nav_date, nav, nav_acc FROM fund_nav where wind_code=%s'
    df = pd.read_sql(sql_str, engine, params=[wind_code])
    df.rename(columns={'nav_date': '日期', 'nav': '净值', 'nav_acc': '累计净值'}, inplace=True)
    df.set_index('日期', inplace=True)
    file_path = get_cache_file_path('%s_nav.xls' % wind_code)
    df.to_excel(file_path, sheet_name=wind_code)


if __name__ == '__main__':
    folder_path = r'D:\WSPycharm\fund_evaluation\contact'
    wind_code = 'fh_0052'
    # import_fund_sec_pct(wind_code, folder_path)
    import_fund_sec_pct(wind_code, folder_path) # , mode=None
    # export_nav_excel(wind_code)
