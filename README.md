# Forecast_by_Leading_Indicators

# -*- coding:utf-8 -*-
import os
import numpy as np
import pandas as pd
import statsmodels.api as sm  # 季调滤波模块

# X11方法的TC季调值，其中输入变量为未季调数据，DaraFrame格式
def TC_X11(before_tc):

    name_list = before_tc.columns.values  # 获得变量名称，第一个为因变量，后面为自变量
    var_num = name_list.shape[0]  # 获取data中的变量个数
    TC_result = pd.DataFrame()  # 存放TC结果

    for i in range(var_num):
        before_tc_var = before_tc.iloc[:, i]  # 拆分出单个变量，用于后续TC
        decomposition = sm.tsa.seasonal_decompose(before_tc_var, model='multiplicative', freq=12,
                                                  extrapolate_trend=1)  # "additive", "multiplicative"
        tc = decomposition.trend  # 单个变量TC后结果

        TC_result[name_list[i] + '_TC'] = tc

    return TC_result

# HP滤波，其中输入变量为TC后的数据，DaraFrame格式
def HP_filter(tc):

    name_list = tc.columns.values  # 获得变量名称，第一个为因变量，后面为自变量
    var_num = name_list.shape[0]  # 获取data中的变量个数
    HP_result = pd.DataFrame()  # 存放滤波结果

    for i in range(var_num):
        tc_var = tc.iloc[:, i]
        cycle, trend = sm.tsa.filters.hpfilter(tc_var, 14400)  # 周期项和趋势项，lamda=14400

        HP_result[name_list[i] + '_cycle'] = cycle
        HP_result[name_list[i] + '_trend'] = trend

    return HP_result

## 主函数
def main():

    # 读取原始数据,要求原始数据是齐的
    parameter_sheet = pd.read_excel('Initial_Data.xlsx', sheet_name='参数设置')  # 读取当前路径下Data，获得参数信息
    data_sheet = pd.read_excel('Initial_Data.xlsx', sheet_name='原始数据', index_col=0)  # 读取当前路径下Data，获得数据表，DataFrame格式

    # 传入参数
    select_TC = parameter_sheet.iloc[2, 1]
    select_HP = parameter_sheet.iloc[3, 1]

    # 执行
    result_1 = pd.DataFrame()
    if select_TC == 1:
        print('步骤1：执行TC（X11）')
        if select_HP == 1:
            print('步骤2：执行HP滤波')
            data_TC = TC_X11(data_sheet)  # TC
            data_HP = HP_filter(data_TC)  # HP滤波
            result_1 = pd.concat([data_sheet, data_TC, data_HP], axis=1, join_axes=[data_sheet.index])  # 合并原始数据和处理后数据
            print('成功输出TC和HP滤波结果')
        else:
            data_TC = TC_X11(data_sheet)  # TC
            result_1 = pd.concat([data_sheet, data_TC], axis=1, join_axes=[data_sheet.index])  # 合并原始数据和处理后数据
            print('成功输出TC结果')
    elif select_TC == 0:
        if select_HP == 1:
            print('步骤1：执行HP滤波')
            data_HP = HP_filter(data_sheet)  # HP滤波
            result_1 = pd.concat([data_sheet, data_HP], axis=1, join_axes=[data_sheet.index])  # 合并原始数据和处理后数据
            print('成功输出HP结果')
        else:
            print('请检查参数设置')

    writer = pd.ExcelWriter('result_1.xlsx')  #用pd.ExcelWriter的好处是，可以将多个DataFrame同时写入不同的Sheet中，而不会被覆盖
    result_1.to_excel(writer)  #若写入多个sheet时，可使用sheet_name="a1"的参数
    writer.save()

if __name__ == "__main__":
    main()
