"""
Created on 2017/4/5
@author: yby
"""
from datetime import datetime
import pandas as pd
import xlrd
import re
import os


# #获取表头
# file_path=r'D:\WSPycharm\fund_evaluation\contact\{}'.format(file_name)
# data_df = pd.read_excel(file_path, skiprows=1, header=0)
# #获取净值
# data_df1 = pd.read_excel(file_path, skiprows=3, header=0)
# cum_nva=data_df1['科目名称'][data_df1['科目代码']=='累计单位净值:']
#
#
#
# ct=dict(date=data_df.iloc[:,0][0][5:],name=data_df.columns[0][13:-6],nva=cum_nva.values[0])
#
# def get_fund_nav_date(file_path):
#     data_xls = xlrd.open_workbook(file_path)  # 打开xls文件
#     # sheet_name = data.sheet_names()[0]
#     data_sheet = data_xls.sheet_by_index(0)
#     val_cell = data_sheet.cell(2, 0).value
#     if val_cell.find('估值日期：') != 0:
#         raise('估值日期未找到：%s' % (file_path))
#     nav_date = datetime.strptime(val_cell[5:], '%Y-%m-%d').date()
#     return nav_date

fund_dictionay={}
folder_path=r'D:\WSPycharm\fund_evaluation\evaluation_table'
file_names = os.listdir(folder_path)
for file_name in file_names:
    # file_path = r'd:\Works\F复华投资\合同、协议\丰润\丰润一期\SK8992_复华丰润稳健一期_估值表_20170113.xls'
    file_path = os.path.join(folder_path, file_name)
    file_name_net, file_extension = os.path.splitext(file_path)
    if file_extension not in ('.xls', '.xlsx') :
        continue
    else:
        file_path='D:\\WSPycharm\\fund_evaluation\\evaluation_table\\万霁九号私募投资基金_估值表.xlsx'
        # file_path = 'D:\\WSPycharm\\fund_evaluation\\evaluation_table\\复华定增1号-2018-08-31.xls'
        data_df = pd.read_excel(file_path, skiprows=1, header=0).dropna(how='all',axis=0).dropna(how='all',axis=1)
        date = try_2_date(data_df.iloc[0][0])  # data_df.iloc[0][0][-10:]
        #获取净值
        data_df1 = pd.read_excel(file_path, skiprows=3, header=0)
        cum_nva=data_df1['科目名称'][data_df1['科目代码']=='累计单位净值:']
        if '财通' in data_df.columns[0]:
            name,nva= data_df.columns[0][13:-6],cum_nva.values[0]
        elif '万霁' in data_df.columns[0]:
            name, nva = data_df.columns[0], cum_nva.values[0]
        # fund_dictionay[name]=nva
        fund_dictionay.setdefault(name, []).append([date, nav])
        # ct=dict(date=data_df.iloc[:,0][0][5:],name=data_df.columns[0][13:-6],nva=cum_nva.values[0])



user = 'baoyuan.yang@foriseinvest.com'
password = 'ybychem87'
imap_url = 'imap.263.net'
#Where you want your attachments to be saved (ensure this directory exists)
attachment_dir = r'D:\WSPycharm\fund_evaluation\Zhanhong'
# sets up the auth
def auth(user,password,imap_url):
    con = imaplib.IMAP4_SSL(imap_url)
    con.login(user,password)
    return con
# extracts the body from the email
def get_body(msg):
    if msg.is_multipart():
        return get_body(msg.get_payload(0))
    else:
        return msg.get_payload(None,True)
# allows you to download attachments
def get_attachments(msg):
    for part in msg.walk():
        if part.get_content_maintype()=='multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        fileName = part.get_filename()

        if bool(fileName):
            filePath = os.path.join(attachment_dir, fileName)
            with open(filePath,'wb') as f:
                f.write(part.get_payload(decode=True))
#search for a particular email
def search(key,value,con):
    result, data  = con.search(None,key,'"{}"'.format(value))
    return data
#extracts emails from byte array
def get_emails(result_bytes):
    msgs = []
    for num in result_bytes[0].split():
        typ, data = con.fetch(num, '(RFC822)')
        msgs.append(data)
    return msgs

con = auth(user,password,imap_url)
con.select('INBOX')

result, data = con.fetch(b'10','(RFC822)')
raw = email.message_from_bytes(data[0][1])
get_attachments(raw)