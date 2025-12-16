import pandas
import openpyxl
from openpyxl.styles import colors
from openpyxl.styles import Font, Color
from openpyxl.styles.borders import Border, Side
from openpyxl import load_workbook
from openpyxl.chart import BarChart3D, Reference, AreaChart, AreaChart3D, Series, LineChart, LineChart3D
import os
import pdb
from openpyxl.styles import NamedStyle, Font, Border, Side
# import excel2img
import time
import datetime
import xlsxwriter
import math
import numpy
from time import gmtime, strftime


def OD_IBS_SLA_STATUS(input_1, input_2, Output_path):
    print('Processing Start Please Wait!')
    start = time.time()
    print('Processing!')
    input_1_df = pandas.read_excel(input_1, sheet_name='CA Integration')
    input_1_df_ = input_1_df[input_1_df['CA Status'] == 'On air']
    input_1_df_['date'] = input_1_df_['Test Date']
    input_1_df_['Week Number'] = input_1_df_.date.apply(lambda x: x.isocalendar()[1])
    del input_1_df_['date']
    weeknum = list(set(input_1_df_['Week Number']))
    onair_sites = []
    for i in range(len(weeknum)):
        output__ = input_1_df_[input_1_df_['Week Number'] <= weeknum[i]]
        onair_sites.append(len(output__))
    OP_CA = pandas.DataFrame()
    OP_CA['Week Number'] = weeknum
    OP_CA['CA On-Air Sites'] = onair_sites

    #========================TXN=============================

    input_TXN_df = pandas.read_excel(str(os.getcwd())+'\\Input\\CA Sites TXN Check.xlsx', sheet_name='Sheet1')
    input_TXN_df_ = input_TXN_df[input_TXN_df['CA status'] == 'On air']
    input_TXN_df_['date'] = input_TXN_df_['Date']
    input_TXN_df_['Week Number'] = input_TXN_df_.date.apply(lambda x: x.isocalendar()[1])
    del input_TXN_df_['date']
    weeknum = list(set(input_TXN_df_['Week Number']))
    onair_sites = []
    for i in range(len(weeknum)):
        output__ = input_TXN_df_[input_TXN_df_['Week Number'] <= weeknum[i]]
        onair_sites.append(len(output__))
    OP_TXN = pandas.DataFrame()
    OP_TXN['Week Number'] = weeknum
    OP_TXN['TXN On-Air Sites'] = onair_sites

    #========================================================


    input_2_df = pandas.read_excel(input_2, sheet_name = '5G RollOut Tracker')
    input_2_df_old = input_2_df[input_2_df['On-Air Date'] <= '2022-01-01']
    input_2_df_22 = input_2_df[input_2_df['On-Air Date'] >= '2022-01-01']
    input_2_df_old['Week Number'] = 0
    input_2_df_22['today'] = input_2_df_22['On-Air Date']
    input_2_df_22['Week Number'] = input_2_df_22.today.apply(lambda x: x.isocalendar()[1])
    del input_2_df_22['today']
    output = input_2_df_old.append(input_2_df_22)
    output_ = output[output['On-Air Sites'] == 'On-Air']
    weeknum = list(set(output_['Week Number']))
    onair_sites = []
    for i in range(len(weeknum)):
        output__ = output_[output_['Week Number'] <= weeknum[i]]
        onair_sites.append(len(output__))
    OP_RO = pandas.DataFrame()
    OP_RO['Week Number'] = weeknum
    OP_RO['RO On-Air Sites'] = onair_sites

    print('Output Path: ', Output_path)

    writer = pandas.ExcelWriter(Output_path + '\\' + 'CA & RollOut.xlsx',engine='xlsxwriter')  # Info: output file dir
    OP_CA.to_excel(writer, sheet_name='CA', index=False)
    OP_RO.to_excel(writer, sheet_name='RO', index=False)
    OP_TXN.to_excel(writer, sheet_name='TXN', index=False)
    writer.save()
    writer.close()

    end = time.time()
    Execute_Time = "{:.3f}".format((end - start) / 60)
    print('The Execution Time of this Tool is %s minutes.' % Execute_Time)
    time.sleep(1)
    print('Execution Completed Succcessfully!')
    time.sleep(1)
    print('')
    print('')
    print('---------------Huawei RF Middle East----------------')
    print('---------For Support: Danish Ali(dwx854280)---------')
    print('---------------Contact: 00971508552942--------------')
    time.sleep(3)


