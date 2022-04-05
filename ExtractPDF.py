#!/usr/bin/env python
# -*- coding: UTF-8 -*-
"""
@Project ：StandardProject
@File    ：ExtractPDF.py
@Author  ：最可爱的Colin
@Date    ：2022/3/31 11:20
@Note    ：

"""

import importlib
import os
import re
import sys

import camelot
import pandas as pd
from pygtrans import Translate

# ----- pdfminer引用位置 -----#
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer

# ----- 与Excel交互的包引用位置 -----#

importlib.reload(sys)

# ----- 全局变量 定义文件的输入和输出的名称文件夹 -----#
INPUT_PDF = 'input_pdf'
OUTPUT_PDF = 'output_pdf_table'
INPUT_EXCEL = 'input_excel'


# 获得目标类型文件下的所有该类型文件路径
def Get_Files_Path(dir_path, file_type):
    """
    :param dir_path:文件存放的目录
    :param file_type:文件类型
    :return:返回由[绝对路径,文件名]组成的list
    """

    # 递归获取子目录中的pdf
    def list_dir(file_dir, path_list):
        """
        :param file_dir:目标文件所在目录
        :param path_list:外部传入的用于存储输出目录信息
        :return:
        """

        # 列出当前目录下所有文件和目录
        dir_list = os.listdir(file_dir)

        # 排序 会影响到后面Excel的合并
        dir_list.sort()
        # print(dir_list)

        # 递归获取目录中的文件
        for cur_file in dir_list:
            # 获取文件的绝对路径
            path = os.path.join(file_dir, cur_file)

            # 如果是文件则读取
            if os.path.isfile(path):  # 判断是否是文件还是目录需要用绝对路径
                file = cur_file
                file_name = os.path.splitext(file)[0]  # 输出：file_test
                file_suffix = os.path.splitext(file)[1]  # 输出：.txt

                if file_suffix == '.' + file_type:
                    path_list += [[path, file_name]]

                    # print(file_name, file_suffix)

            # 如果是目录则递归
            if os.path.isdir(path):
                list_dir(path, dir_path)  # 递归子目录

    # 存储所有文件的路径
    file_path = []
    list_dir(dir_path, file_path)

    print('执行分析的文件目录和数量: ', dir_path, len(file_path))

    return file_path


# 在根目录创建一个文件夹用于存储提取的表格
def Create_Dir_of_Tables(dirname):
    """
    :param dirname:存放表格的目录名
    :return: 成功失败提示
    """

    if not os.path.exists(dirname):
        os.makedirs(dirname)
        return dirname
    else:
        return dirname


# 手动输入目标表格页面所在的页码
def Get_Tables_Pages(pdf_name):
    """
    :param pdf_name:传入指定的pdf名称
    :return: 返回表格所在的开始和结束的页码list
    """
    table_pages = {'cafr_2010': [87, 92], 'cafr_2011': [87, 92], 'cafr_2012': [85, 91], 'cafr_2013': [89, 95],
                   'cafr_2014': [111, 118],
                   'cafr_2015': [109, 116], 'cafr_2016': [115, 122], 'cafr_2017': [127, 133],
                   'cafr_2018': [110, 113],
                   'cafr_2019': [110, 114],
                   'cafr_2020': [114, 119], 'cafr_2021': [118, 123]}

    return table_pages[pdf_name]


# 手动输入目标文字页面所在的页码
def Get_Texts_Pages(pdf_name):
    """
       :param pdf_name:传入指定的年份
       :return: 返回投资章节中的概要所在的开始和结束的页码list
       """
    text_pages = {'cafr_2010': [86, 86], 'cafr_2011': [86, 86], 'cafr_2012': [84, 84], 'cafr_2013': [88, 88],
                  'cafr_2014': [109, 110],
                  'cafr_2015': [107, 108], 'cafr_2016': [113, 114], 'cafr_2017': [126, 126],
                  'cafr_2018': [109, 109],
                  'cafr_2019': [109, 109],
                  'cafr_2020': [113, 113], 'cafr_2021': [117, 117]}
    return text_pages[pdf_name]


# 使用Camelot获取单个PDF指定页码内容中的表格
def Extract_Tables(Pages, Pdf):
    """
    :param Pages:需要提取表格的页码
    :param Pdf:pdf的路径和文件名
    :return: 每张PDF输出一张包含了指定页面的提取出表格的Excel文件
    """

    # 分离开始和结束页面
    page_start, page_end = Pages

    # Camelot传入的页码不是以0开始,需要+1
    page_start += 1
    page_end += 1

    # 分离pdf的路径和文件名
    pdf_path, pdf_name = Pdf

    # 一次获得该PDF中所有表格
    # camelot的参数改为stream识别没有外框的表格,edge_tol调整大一些获取更大的识别范围
    # 如果需要提高精确度可以使用调试模式,但是实际结果还不错所以没有可视化调试
    tables = camelot.read_pdf(filepath=pdf_path, pages=str(page_start) + '-' + str(page_end), flavor='stream',
                              edge_tol=1000, flage_size=True)

    # 遍历该PDF中的所有表格并存在一个Excel中
    with pd.ExcelWriter("pdf_tables/" + pdf_name + '.xlsx') as writer:
        for table in tables:
            # df提供了将table转化为DataFrame的方法
            df_table = pd.DataFrame(table.df)

            # 数据清洗
            df_table.replace(['\\(cid:3\\)', '\\(cid:882\\)'], ' ', inplace=True, regex=True)
            # print(df_table)

            # table.parsing_report这个函数可以列出识别报告,包括表格所在页面和顺序用于命名
            table_report = table.parsing_report

            # 把一张PDF里面的所有表格按照多个工作表储存在一个Excel工作薄中
            df_table.to_excel(writer, sheet_name=str(str(table_report['page']) + '-' + str(table_report['order'])))


# 使用Pdfminer获取单个PDF指定页码内容中的文字
def Extract_Texts(Pages, Pdf):
    # 分离pdf的路径和文件名
    pdf_path, pdf_name = Pdf

    client = Translate()
    with open('TEXT_SUM.txt', 'a+') as file_text:

        # 定义一个分隔符用于分割pdf
        pdf_split_flag = "\n\n-------------[" + pdf_name + "]-------------\n\n"

        # 插入分隔符
        print(pdf_split_flag, file=file_text)

        # 插入分隔符
        with open('TEXT_SUM_TRANS.txt', 'a+') as file_trans:
            print(pdf_split_flag, file=file_trans)

        for page_layout in extract_pages(pdf_path, page_numbers=Pages):
            for element in page_layout:
                if isinstance(element, LTTextContainer):
                    text_result = element.get_text()

                    # 数据清洗
                    text_clean = text_result.replace("(cid:3)", ' ')
                    text_clean = re.sub(r"\s+", ' ', text_clean)
                    text_clean = re.sub(r" +", ' ', text_clean)

                    # 数据翻译
                    if text_clean != " ":
                        # print(repr(text_clean),type(text_clean))

                        try:
                            text_trans = client.translate(text_clean)
                            text_trans_text = text_trans.translatedText

                            with open('TEXT_SUM_TRANS.txt', 'a+') as file_trans:
                                print(text_trans_text + '\n', file=file_trans)
                        # 翻译会遇到网络问题导致错误,这个是接口不稳定的问题,多重新运行几遍就行了

                        # 异常处理
                        except AttributeError as ae:
                            print(ae)
                            continue

                    print(text_clean, file=file_text)


# 用于处理excel中的表格合并
def Merge_By_Sheet(excel_path):
    # 所有excel文件所在的目录
    files = Get_Files_Path(excel_path, 'xlsx')
    # print(files)

    # 创建合并后的文件夹
    target_path = Create_Dir_of_Tables(dirname='output_excel' + '_merge')

    # 创建一个Dict用于存储不同的sheet姓名
    dict_sheet = {}

    # 遍历每个excel文件获得名称
    for file_path, file_name in files:
        result = pd.read_excel(file_path, sheet_name=None)

        # 获得result中的sheet获得所有的名称
        for key, value in result.items():
            key = key.strip().upper()
            # print(key)
            dict_sheet.update({key: []})

            # 防止过多的表格出错
            if len(dict_sheet) > 20:
                raise ValueError

    # 获得内容
    for file_path, file_name in files:
        result = pd.read_excel(file_path, sheet_name=None)
        for sheet_name, sheet_value in result.items():
            sheet_name = sheet_name.strip().upper()
            # print(sheet_name)
            dict_sheet[sheet_name] += [[file_name] + [sheet_value]]
    # DICT_Sheet.update({sheet_name.strip(): DICT_Sheet[sheet_name.strip()]+sheet_value})
    # DICT_Sheet[sheet_name.strip()]=DICT_Sheet[sheet_name.strip()]+[sheet_value]
    # print("长度", len(DICT_Sheet[key.strip()]), type(DICT_Sheet[key.strip()]))
    # print(sheet_list)

    # 按照sheet重新组成excel
    for key, value in dict_sheet.items():

        # 遍历所有的sheet,value是一个list(Dataframe)
        excel_writer = pd.ExcelWriter(target_path + '/' + key + '.xlsx', engine='xlsxwriter')

        for sheet in value:
            sheet_name, sheet_content = sheet
            # print(sheet_content)
            # time.sleep(1000)
            df_sheet = pd.DataFrame(sheet_content)
            df_sheet.to_excel(excel_writer, index=False, sheet_name=sheet_name)
            # excel_writer.save()

        excel_writer.save()


# 打包好的用于获取FSA10年PDF
def Packge_FSAPDF10():
    # 获取指定目录中所有要解析的pdf路径和文件名
    pdfs = Get_Files_Path(INPUT_PDF, 'pdf')

    # 创建一个文件夹用于储存表格
    Create_Dir_of_Tables(OUTPUT_PDF)

    # 依次解析所有的pdf
    for PDF in pdfs:
        # 分离pdf的路径和文件名
        pdfpath, pdfname = PDF

        print("正在提取", pdfpath)

        # ----------------- 0.Camelot提取表格 ----------------- #

        # 获得PDF文件中表格所在的页码范围
        tables_pages = Get_Tables_Pages(pdfname)
        # 提取
        Extract_Tables(tables_pages, PDF)

        # ----------------- 1.Pdfminer提取文本 ----------------- #

        # # 获得PDF文件中文本所在的页码范围
        # Texts_Pages = Get_Texts_Pages(pdfname)
        # # 提取
        # Extract_Texts(Texts_Pages, PDF)


if __name__ == '__main__':
    # Merge_By_Sheet('tables_2010_2016')
    # Packge_FSAPDF10()
    Merge_By_Sheet(INPUT_EXCEL)
    # python ExtractPDF.py
