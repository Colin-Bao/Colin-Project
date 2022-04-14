#!/usr/bin/env python
# -*- coding: UTF-8 -*-
"""
@Project ：StandardProject
@File    ：ExtractPDF.py
@Author  ：最可爱的Colin
@Date    ：2022/3/31 11:20
@Note    ：Excel各种变换

"""
import pandas as pd


# 用于把一行的多个列属性变成行，如[A:{[最大值:1,最小值:0]}]变成[{A最大值:1}，{A最小值:0}]

def table_transform(excel_path, excel_sheet):
    df_excel = pd.read_excel(excel_path, sheet_name=excel_sheet)

    # 新建一个df
    df_new = pd.DataFrame(columns=('Year', 'Class', 'Class-2', 'Value'))
    for index, row in df_excel.iterrows():
        # print(row)

        # 数据清洗
        row['Class'] = row['Class'].strip()

        if 'TOTAL' in row['Class'].upper():
            row['Class'] = row['Class'].strip().upper()

        for key, value in row.items():

            if key != 'Year' and key != 'Class':
                # 把Value转换为list对象再传入才能调用concat方法
                add_dict = {'Year': [row['Year']], 'Class': [row['Class']], 'Class-2': [key],
                            'Value': [value]}
                add_df = pd.DataFrame(add_dict)
                df_new = pd.concat([df_new, add_df], ignore_index=True)
    # 数据清洗
    df_new.replace(
        ['US Treasuries', 'USA', 'Non-U.S.', 'TRS ', 'Global Inflation Linked Bonds', 'TIPS',
         'Real Estate', 'Non-US'],
        ['U.S. Treasuries', 'U.S.', 'Non U.S.', '', 'Inflation Linked Bonds',
         'Inflation Linked Bonds', 'Real Assets', 'Non U.S.'],
        inplace=True, regex=True)
    df_new.replace(['[\\s]+'], [' '],
                   inplace=True, regex=True)
    df_new.replace(
        ['Real Estate', 'Government Bonds', 'Energy, Natural Resources, and Infrastructure', 'TOTAL PUBLIC EQUITY',
         'Fixed Income', 'Stable Value Hedge Funds'],
        ['Real Assets', 'U.S. Treasuries', 'Energy, Natural Resources & Infrastructure', 'Public Equity',
         'Absolute Return', 'Hedge Funds'],
        inplace=True, regex=True)
    df_new.to_excel('Policy.xlsx')


table_transform('POLICY RANGES.xlsx', 'Merge')
