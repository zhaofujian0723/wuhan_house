# -*- coding: utf-8 -*-

import requests
from lxml import etree
import xlrd
import xlwt
from xlutils.copy import copy
import os


def write_excel_xls(path, sheet_name, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlwt.Workbook()  # 新建一个工作簿
    sheet = workbook.add_sheet(sheet_name)  # 在工作簿中新建一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            sheet.write(i, j, value[i][j])  # 像表格中写入数据（对应的行和列）
    workbook.save(path)  # 保存工作簿


def write_excel_xls_append(path, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            new_worksheet.write(i+rows_old, j, value[i][j])  # 追加写入数据，注意是从i+rows_old行开始写入
    new_workbook.save(path)  # 保存工作簿


def get_page(url, page):
    FramData = {
        "__VIEWSTATE": "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",
        "__EVENTTARGET": "AspNetPager1",
        "__EVENTARGUMENT": "{}".format(page),
        "__EVENTVALIDATION": "/wEdABchLbp5aMmc7vqD32Zbz8eI6qGgHJC3rzvP9cQgZz7WCeZ5PN9t4A7d5BkgEXZCI2OsNuh1Du9QvCb+KLdY9QD5w0lzwUrl67DQCd422Ua9fkYr1bL4Zhh9nXfDy6F2tlMbWlbQa8axP65+97xBEc1k1qImOu8X2+R9mV6kg7sdnXvSD0gEBuIGxfsGe35GnMpyOoKvQny89caM/OF84k+H/yqLxnZIhS09L+HMOXPEE6pk7zOPM2jeC9LnUDyP0+rY25Fr/QL6eB7VI9vx0mEgaKO8Hdm8T5ZM+WcyyRWeWWumNZdrbu21Hbq2YT58cCdxJs7jfQpho4V2byUyfJxKeGy8LbC16/BeqQ1V7cVtpX7xGAef0fNuthECH7xGkKpGLy0G2B7hUeW3YwkXEgb98MbpsmWVLCPCqLaU7l4JbCFHOSo/jgtppFZ/rlsSehLp0n5Fjwr5uJwKpbkwfd/b0/KdySeuRpi7LRO6xG0pphBU1Je7Qsul2ICBZGX7uX5Tmq0rCgYI+qs+OwuGfcmO",
        "tbtxt": "(unable to decode value)",
        "DropDownList_xzq": "",
        "xmmc": "",
        "xmdz": "",
        "kfs": "",
        "AspNetPager1_input": "{}".format(page),
    }
    content = requests.post(url, data=FramData).text
    html = etree.HTML(content)
    housing_list = html.xpath('//table[@id="tables"]/tr')
    for h in housing_list:
        housing_name = h.xpath('./td[1]/a/text()')
        if housing_name:
            housing_name = housing_name[0]  # 项目名称
            housing_num = h.xpath('./td[2]/text()')[0]  # 总套数(套)
            sold_housing_num = h.xpath('./td[3]/text()')[0]  # 住房已售(套)
            available_housing_num = h.xpath('./td[4]/text()')[0]  # 住房可售(套)
            sold_not_housing_num = h.xpath('./td[5]/text()')[0]  # 非住房已售(套)
            available_not_housing_num = h.xpath('./td[6]/text()')[0]  # 非住房可售(套)
            print(housing_name, housing_num)
            write_excel_xls_append(path_name, [[housing_name, housing_num, sold_housing_num, available_housing_num, sold_not_housing_num, available_not_housing_num]])
    if len(housing_list) >= 2:
        page += 1
        if 280 >= page:
            get_page(url, page)


if __name__ == '__main__':
    link = " http://119.97.201.22:8083/search/spfxmcx/spfcx_index.aspx"
    path_name = '住房.xls'
    sheet_ = '武汉'
    s_list = []
    if not os.path.exists(path_name):
        title_list = [["项目名称", "总套数(套)", "住房已售(套)", "住房可售(套)", "非住房已售(套)", "非住房可售(套)"]]
        write_excel_xls(path_name, sheet_, title_list)
    else:
        os.remove(path_name)
        title_list = [["项目名称", "总套数(套)", "住房已售(套)", "住房可售(套)", "非住房已售(套)", "非住房可售(套)"]]
        write_excel_xls(path_name, sheet_, title_list)
    get_page(link, 1)
