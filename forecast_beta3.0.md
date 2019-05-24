# 快捷键 撤销：Ctrl+Z  反撤销：Ctrl+Shift+Z  注释/反注释：Ctrl+/
# 一定注意相对路径的写法，用os调用R，打开文件的路径要写相对于这个python的路径
# result_code = os.system("Rscript X12/TC.R ") #os.system(Rscript的路径 R文件的路径，参数)，其中参数只能是字符串，比如说路径，行业的字符串等等

# -*- coding:utf-8 -*-
import os
import time
import threading
# ----------------------------------------------------------------------------------------
import numpy as np
import pandas as pd
import statsmodels.api as sm  # 计量模块，包括季调滤波方法，和线性回归方法
# ----------------------------------------------------------------------------------------
import tkinter as tk
# ----------------------------------------------------------------------------------------
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg  # 将matplotlib图像绘制在tkinter的Canvas上的模块

###############################################################功能1##############################################################
# X11方法的TC季调值，其中输入变量为未季调数据，DaraFrame格式
def tc_x11(before_tc):

    name_list = before_tc.columns.values  # 获得变量名称，第一个为因变量，后面为自变量
    var_num = name_list.shape[0]  # 获取data中的变量个数
    tc_result = pd.DataFrame()  # 存放TC结果

    for i in range(var_num):
        before_tc_var = before_tc.iloc[:, i]  # 拆分出单个变量，用于后续TC
        decomposition = sm.tsa.seasonal_decompose(before_tc_var, model='multiplicative', freq=12,
                                                  extrapolate_trend=1)  # "additive", "multiplicative"
        tc = decomposition.trend  # 单个变量TC后结果

        tc_result[name_list[i] + '_TC'] = tc

    return tc_result

# X12方法的TC季调值，其中输入变量为未季调数据，DaraFrame格式
def tc_x12(before_tc):

    # 把等待X12季调的数据输出到excel，以备R读取
    writer = pd.ExcelWriter('result_temp1.xlsx')  # 用pd.ExcelWriter的好处是，可以将多个DataFrame同时写入不同的Sheet中，而不会被覆盖
    before_tc.to_excel(writer)  # 若写入多个sheet时，可使用sheet_name="a1"的参数
    writer.save()

    # 调用R程序，os_system(Rscript的路径 R文件的路径，参数)，其中，参数只能是字符串，比如说路径，行业的字符串等等
    os.system("Rscript X12.R ")
    tc_result = pd.read_excel('result_temp2.xlsx', index_col=0)  # 读取利用R做出的TC结果excel，DataFrame格式

    # 删除过程文件
    os.remove('result_temp1.xlsx')
    os.remove('result_temp2.xlsx')

    return tc_result

# HP滤波，其中输入变量为TC后的数据，DaraFrame格式
def hp_filter(tc):

    name_list = tc.columns.values  # 获得变量名称，第一个为因变量，后面为自变量
    var_num = name_list.shape[0]  # 获取data中的变量个数
    hp_result = pd.DataFrame()  # 存放滤波结果

    for i in range(var_num):
        tc_var = tc.iloc[:, i]
        cycle, trend = sm.tsa.filters.hpfilter(tc_var, 14400)  # 周期项和趋势项，HP滤波模型中参数lamda=14400

        hp_result[name_list[i] + '_cycle'] = cycle
        hp_result[name_list[i] + '_trend'] = trend

    return hp_result

# 功能1执行函数，目前TC功能执行的是X12
def execute_function_1(data_sheet, parameter_sheet):

    # 传入参数信息，行从-1开始数，列从0开始数
    select_tc = parameter_sheet.iloc[4, 1]  # TC选择
    select_hp = parameter_sheet.iloc[5, 1]  # HP滤波选择

    result_1 = pd.DataFrame()  # 创建空DataFrame，用于后续保存结果

    # 根据参数执行功能
    if select_tc == 1:
        print('步骤1：执行TC（X12）')
        data_tc = tc_x12(data_sheet)  # TC
        if select_hp == 1:
            print('步骤2：执行HP滤波')
            data_hp = hp_filter(data_tc)  # HP滤波
            result_1 = pd.concat([data_sheet, data_tc, data_hp], axis=1, join_axes=[data_sheet.index])  # 合并原始数据和处理后数据
            print('成功输出TC和HP滤波结果')
        else:
            result_1 = pd.concat([data_sheet, data_tc], axis=1, join_axes=[data_sheet.index])  # 合并原始数据和处理后数据
            print('成功输出TC结果')
    elif select_tc == 0:
        if select_hp == 1:
            print('步骤1：执行HP滤波')
            data_hp = hp_filter(data_sheet)  # HP滤波
            result_1 = pd.concat([data_sheet, data_hp], axis=1, join_axes=[data_sheet.index])  # 合并原始数据和处理后数据
            print('成功输出HP滤波结果')
        else:
            print('请检查功能1的参数设置')

    return result_1

###############################################################功能2##############################################################
# step1里按钮命令，待预测指标按钮，注意！按钮里不能传入参数
def selection_lag():

    # 直接通过global将data_show传进来
    global listbox_var
    global name_lag
    name_lag = listbox_var.get(listbox_var.curselection())  # 获取Listbox控件中光标选择的值
    global label_name_lag_text
    label_name_lag_text.set('待预测指标：' + name_lag)  # 将选中的名称传到Label上

    # 操作fig_1，调出待预测指标（name_lag）对应的数据并放在图上
    global data_show
    global fig_1
    global line_lag

    # 打开图形交互模式，从而反复操作fig_1
    plt.ion()

    # 判断line_lag之前是否绘制，若存在，则删除之前的line_lag并绘制新的：若不存在，则绘制
    try:
        line_lag
    except NameError:
        pass
    else:
        fig_1.lines.remove(line_lag[0])  # 删除之前的待预测指标图形

    date = data_show['日期']  # x轴，日期
    data = data_show[name_lag]  # y轴

    line_lag = fig_1.plot(date, data, color='blue')  # 绘图

    # 自动调整y坐标轴最大值和最小值，最大值等于name_lag和name_lead（若存在）的最大值，最小值同理
    global name_lead
    try:
        line_lead
    except NameError:
        data_max = max(data_show[name_lag])
        data_min = min(data_show[name_lag])
    else:
        data_max = max(max(data_show[name_lag]), max(data_show[name_lead]))
        data_min = min(min(data_show[name_lag]), min(data_show[name_lead]))

    plt.ylim(data_min, data_max)

    # 关闭图形交互模式
    plt.ioff()

    global canvas
    canvas.draw()

# step1里按钮命令，领先指标按钮
def selection_lead():

    global listbox_var
    global name_lead
    name_lead = listbox_var.get(listbox_var.curselection())  # 获取Listbox控件中光标选择的值
    global label_name_lead_text
    label_name_lead_text.set('领先指标：' + name_lead)  # 将选中的名称传到Label上

    # 操作fig_1，调出待预测指标（name_lead）对应的数据并放在图上
    global data_show
    global fig_1
    global line_lead

    # 打开图形交互模式，从而反复操作fig_1
    plt.ion()

    # 判断line_lead之前是否绘制，若存在，则删除之前的line_lead并绘制新的：若不存在，则绘制
    try:
        line_lead
    except NameError:
        pass
    else:
        fig_1.lines.remove(line_lead[0])  # 删除之前的待预测指标图形

    date = data_show['日期']  # x轴，日期
    data = data_show[name_lead]  # y轴

    line_lead = fig_1.plot(date, data, color='red')  # 绘图

    # 自动调整y坐标轴最大值和最小值，最大值等于name_lag（若存在）和name_lead的最大值，最小值同理
    global name_lag
    try:
        line_lag
    except NameError:
        data_max = max(data_show[name_lead])
        data_min = min(data_show[name_lead])
    else:
        data_max = max(max(data_show[name_lag]), max(data_show[name_lead]))
        data_min = min(min(data_show[name_lag]), min(data_show[name_lead]))

    plt.ylim(data_min, data_max)

    # 关闭图形交互模式
    plt.ioff()

    # 传递用于移动的领先指标数据
    global data_lead_move
    data_lead_move = data_show[name_lead]

    # 由于选择了新的领先指标，将之前的领先期数重置为零
    global var_lead_num
    var_lead_num = 0
    global label_lead_text
    label_lead_text.set(var_lead_num)

    global canvas
    canvas.draw()

# step2里按钮命令，领先指标右移
def right_move():

    global var_lead_num
    var_lead_num += 1
    global label_lead_text
    label_lead_text.set(var_lead_num)

    # 实现fig_1中line_lead的右移，提取name_lead列，shift一下，然后画图
    global name_lead
    global line_lead

    global data_show
    global data_lead_move

    # 打开图形交互模式，从而反复操作fig_1
    plt.ion()

    fig_1.lines.remove(line_lead[0])  # 删除之前的待预测指标图形

    date = data_show['日期']  # x轴，日期
    data_lead_move = data_lead_move.shift(1)  # y轴

    line_lead = fig_1.plot(date, data_lead_move, color='red')  # 绘图

    # 关闭图形交互模式
    plt.ioff()

    global canvas
    canvas.draw()

# step2里按钮命令，领先指标左移
def left_move():

    # 实现fig_1中line_lead的左移
    # 左移逻辑：提取line_lead数据，并对其第一行不断删除NA，从而实现左移，若第一行没有NA，则不操作
    global data_lead_move

    if pd.isna(data_lead_move[0]) is True:
        # 领先期计数-1
        global var_lead_num
        var_lead_num -= 1
        global label_lead_text
        label_lead_text.set(var_lead_num)

        # 打开图形交互模式，从而反复操作fig_1
        plt.ion()

        global line_lead
        fig_1.lines.remove(line_lead[0])  # 删除之前的待预测指标图形

        global data_show
        date = data_show['日期']  # x轴，日期
        data_lead_move = data_lead_move.shift(-1)  # y轴

        line_lead = fig_1.plot(date, data_lead_move, color='red')  # 绘图

        # 关闭图形交互模式
        plt.ioff()

        global canvas
        canvas.draw()

# 功能2：展示功能，GUI
def execute_function_2():

    ## 创建窗体
    root = tk.Tk()  # 创建窗体
    root.title('待预测数据展示（TJD腾景数研）')  # 窗体名称
    root.geometry('1280x720')  # 窗体大小

    # Frame控件用于分割母窗体root
    frm_left = tk.Frame(root)
    frm_left.pack(side='left')
    frm_right = tk.Frame(root)
    frm_right.pack(side='right')

    frm_left_top = tk.Frame(frm_left)
    frm_left_top.pack(side='top', fill=tk.BOTH)
    frm_left_bottom = tk.Frame(frm_left)
    frm_left_bottom.pack(side='bottom')

    frm_left_bottom_left = tk.Frame(frm_left_bottom)
    frm_left_bottom_left.pack(side='left')
    frm_left_bottom_right = tk.Frame(frm_left_bottom)
    frm_left_bottom_right.pack(side='right')


    ## step1:选择待预测指标和领先指标
    # 标签，显示待预测和领先指标名称
    global label_name_lag_text
    label_name_lag_text = tk.StringVar()
    label_name_lag = tk.Label(frm_left_bottom_left, textvariable=label_name_lag_text,
                              bg='SkyBlue', font=('Arial', 12), width=40, height=3)
    label_name_lag.grid(row=1, column=1)

    global label_name_lead_text
    label_name_lead_text = tk.StringVar()
    label_name_lead = tk.Label(frm_left_bottom_left, textvariable=label_name_lead_text,
                               bg='MistyRose', font=('Arial', 12), width=40, height=3)
    label_name_lead.grid(row=1, column=2)

    # 按钮，用于选定待预测和领先指标，并调出数据做图
    b1 = tk.Button(frm_left_bottom_left, text='待预测指标', font=('宋体', 15),
                   width=10, height=2, command=selection_lag)  # 注意！这里的按钮函数不能加（）
    b1.grid(row=2, column=1)

    b2 = tk.Button(frm_left_bottom_left, text='领先指标', font=('宋体', 15),
                   width=10, height=2, command=selection_lead)
    b2.grid(row=2, column=2)

    # 滚动条，用于滚动列表框
    scorllbar = tk.Scrollbar(frm_right)
    scorllbar.pack(side='right', fill=tk.Y)

    # 列表框，存放待预测指标名称
    global data_show  # 直接通过global将data_show传进来
    name_list = list(data_show.columns.values)[1:]  # 获得变量名称，并将array格式转换为list

    var_name = tk.StringVar()
    var_name.set(name_list)  # 指标名称传入var_name
    global listbox_var
    listbox_var = tk.Listbox(frm_right, listvariable=var_name, font=('Arial', 15),
                             yscrollcommand=scorllbar.set, width=40, height=30)  # 领先指标列表，var_lead_num移动这个里面的变量
    listbox_var.pack(fill=tk.BOTH)

    # 关联listbox的y方向view，使滚动条生效
    scorllbar.config(command=listbox_var.yview)

    # 画布，并将matplotlib绘制的图存放进来
    f_1 = plt.figure(figsize=(10, 5))  # 创建绘制图像
    global fig_1
    fig_1 = plt.subplot(1, 1, 1)

    global canvas
    canvas = FigureCanvasTkAgg(f_1, frm_left_top)  # 把绘制图像显示到tkinter窗口上
    canvas.draw()  # canvas上显示，代替plt.show()
    canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)  # 设置canvas位置


    ## step2:移动领先指标
    global var_lead_num  # 领先指标的领先期数，只要在函数里面的变量全局使用都需要global
    var_lead_num = 0
    global label_lead_text
    label_lead_text = tk.StringVar()  # label里面的显示文本控件

    label_lead = tk.Label(frm_left_bottom_right, textvariable=label_lead_text,
                          bg='MistyRose', font=('Arial', 15), width=6, height=2)
    label_lead.grid(row=1, column=2)

    # 按钮，领先指标左移
    button_reduce = tk.Button(frm_left_bottom_right, text='左移', font=('宋体', 15),
                              width=10, height=2, command=left_move)
    button_reduce.grid(row=2, column=1)

    # 按钮，领先指标右移
    button_add = tk.Button(frm_left_bottom_right, text='右移', font=('宋体', 15),
                           width=10, height=2, command=right_move)
    button_add.grid(row=2, column=3)


    ## 进入消息循环
    root.mainloop()

###############################################################功能3##############################################################
# 预测回测的核心回归模块，李靖恒写的
def forecast_core(data_history, data_forecast):

    data_h = pd.DataFrame(data_history)
    data_f = pd.DataFrame(data_forecast)

    ####### regression using history data ######

    Y = data_h.iloc[:, 0]
    x_v = data_h.iloc[:, 1:]

    Y_name = data_h.columns[0]

    x_n = x_v.shape[1]

    X = sm.add_constant(x_v)

    model = sm.OLS(Y, X)
    results = model.fit()

    para = results.params
    rsqured = results.rsquared

    ###### forecasting #######

    X_f = sm.add_constant(data_f)

    X_fm = np.matrix(X_f)
    par_m = np.matrix(para)

    par_m = np.transpose(par_m)

    Y_f = X_fm * par_m

    Y_f = pd.DataFrame(Y_f)

    Y_f.index = data_f.index.values
    Y_f.columns = ['forecast_data']

    ##  Y_f.to_excel("forecast_results.xlsx")

    regression_funcion = Y_name + "=" + str(para[0])

    i = 1
    while i <= x_n:
        regression_funcion = '%s%+f%s%s%s%s' % (regression_funcion, para[i], '*', X.columns[i], '(', ')')
        i += 1

    rsqured = round(rsqured, 3)

    regression_r = [regression_funcion, rsqured]

    Y_f = pd.DataFrame(Y_f)

    return Y_f, regression_r

# 预测模块，传入的三个参数类型分别是DataFrame，List，int
def forecast(data_to_use, lead_num_list, forecast_sample_num):

    ##################################完成后根据情况封装成函数###########################
    # 浅拷贝，以防因指针问题导致data_to_use被修改
    data_to_forecast = data_to_use.copy()

    # 将自变量根据领先期数向下移动，以便后续直接切片回归
    for i in range(len(lead_num_list)):
        data_to_forecast.iloc[:, i+1] = data_to_forecast.iloc[:, i+1].shift(lead_num_list[i])

    # 切除移动后，开头数据不齐的行
    drop_num = max(lead_num_list)  # 切除期数，等于自变量中的最大领先期数
    data_to_forecast = data_to_forecast.iloc[drop_num:, :]  # 切除开头

    # 预测期数，等于自变量中的最小领先期数
    forecast_num = min(lead_num_list)

    # 切出可用于拟合的全部集合，并计算其历史期数
    history_set = data_to_forecast.dropna(how='any')
    history_num = history_set.iloc[:, 0].size

    # 切出可用于预测的全部集合
    forecast_set = data_to_forecast.iloc[history_num:history_num + forecast_num, 1:]
    ###############################################################################

    # 根据forecast_sample_num得到实际用于拟合的集合，forecast_sample_num==0的时，使用全样本回归
    if forecast_sample_num == 0:
        regression_set = history_set
    else:
        regression_set = history_set.iloc[history_num-forecast_sample_num:, :]

    # 进行预测，返回预测结果，以及包含方程和R方的说明文件
    result_3_forecast, result_3_forecast_info = forecast_core(regression_set, forecast_set)

    # 修改列名，从默认的forecast_data改为forecast
    result_3_forecast.rename(columns={'forecast_data': 'forecast'}, inplace=True)

    print(result_3_forecast_info)

    return result_3_forecast


# 回测模块，传入的三个参数类型分别是DataFrame，List，int
def back_test(data_to_use, lead_num_list, back_test_sample_num):

    ##################################完成后根据情况封装成函数###########################
    # 浅拷贝，以防因指针问题导致data_to_use被修改
    data_to_back_test = data_to_use.copy()



    # 将自变量根据领先期数向下移动，以便后续直接切片回归
    for i in range(len(lead_num_list)):
        data_to_back_test.iloc[:, i+1] = data_to_back_test.iloc[:, i+1].shift(lead_num_list[i])

    # 切除移动后，开头数据不齐的行
    drop_num = max(lead_num_list)  # 切除期数，等于自变量中的最大领先期数
    data_to_back_test = data_to_back_test.iloc[drop_num:, :]  # 切除开头

    # 预测期数，等于自变量中的最小领先期数
    forecast_num = min(lead_num_list)

    # 切出可用于拟合的全部集合，并计算其历史期数
    history_set = data_to_back_test.dropna(how='any')
    history_num = history_set.iloc[:, 0].size

    # 切出可用于预测的全部集合
    forecast_set = data_to_back_test.iloc[history_num:history_num + forecast_num, 1:]

    ###############################################################################

    # 思路:合并history_set和forecast_set，根据训练集从头逐一进行拟合，并根据预测期带入数据回测，
    # 直到回测第一个数据为history_num时停止回测，并输出回测结果的DataFrame（平行四边形）

    back_test_set = pd.concat([history_set, forecast_set], join_axes=[history_set.columns])

    # 创建存放全部回测结果的DataFrame
    result_3_back_test = pd.DataFrame(index=data_to_use.index)

    # 进行回测，总共的拟合方程数为history_num-back_test_sample_num+1
    for i in range(history_num-back_test_sample_num+1):
        # 训练集和测试集
        train_set = back_test_set.iloc[i:back_test_sample_num+i, :]
        test_set = back_test_set.iloc[back_test_sample_num+i:back_test_sample_num+i+forecast_num, 1:]

        # 进行回测，返回回测结果，以及包含方程和R方的说明文件（一期的回测结果，后面需要整合）
        result_3_back_test_temp, result_3_back_test_info_temp = forecast_core(train_set, test_set)

        # 修改列名，从默认的forecast_data改为back_test
        result_3_back_test_temp.rename(columns={'forecast_data': 'back_test'}, inplace=True)

        # 整合回测结果
        result_3_back_test = pd.concat([result_3_back_test, result_3_back_test_temp], axis=1, join_axes=[result_3_back_test.index])

    return result_3_back_test


def selection_forecast():

    global listbox_forecast_var
    global name_forecast
    name_forecast = listbox_forecast_var.get(listbox_forecast_var.curselection())  # 获取Listbox控件中光标选择的值

    # 打开Listbox选中名字对应的数据，DataFrame，第一列是日期列
    global forecast_to_use
    forecast_to_use = forecast_result.parse(sheet_name=name_forecast)

    # 操作fig_2，调出待指标（name_forecast）对应的历史数据和预测数据，并绘制在图上
    global fig_2
    global line_history  # 历史数据用这个变量
    # 预测和回测均用这个变量
    global line_forecast  # 这里用于预测

    # 打开图形交互模式，从而反复操作fig_2
    plt.ion()

    # 判断line_history之前是否绘制，若存在，则删除之前的并绘制新的：若不存在，则绘制
    try:
        line_history
    except NameError:
        pass
    else:
        fig_2.lines.remove(line_history[0])  # 删除之前的待预测指标图形

    date = forecast_to_use['日期']  # x轴，日期
    data_history = forecast_to_use.iloc[:, 1]  # y轴，历史数据

    line_history = fig_2.plot(date, data_history, color='blue')  # 绘图

    # 判断line_forecast之前是否绘制，若存在，则删除之前的并绘制新的：若不存在，则绘制
    try:
        line_forecast
    except NameError:
        pass
    else:
        fig_2.lines.remove(line_forecast[0])  # 删除之前的待预测指标图形

    date = forecast_to_use['日期']  # x轴，日期
    data_forecast = forecast_to_use.iloc[:, 2]  # y轴，历史数据

    line_forecast = fig_2.plot(date, data_forecast, color='red')  # 绘图

    # 自动调整y坐标轴最大值和最小值，最大值等于line_history和line_forecast的最大值，最小值同理
    data_max = max(max(data_history), max(data_forecast))
    data_min = min(min(data_history), min(data_forecast))

    plt.ylim(data_min, data_max)

    # 关闭图形交互模式
    plt.ioff()

    # 预测上或下移次数
    global forecast_move_num
    forecast_move_num = 0

    # 存放上下移动后的结果
    global data_forecast_move
    data_forecast_move = data_forecast

    global canvas_2
    canvas_2.draw()

# 实现fig_2中line_forecast的上移，然后画图
def move_up():

    global forecast_move_num
    forecast_move_num += 1

    # 打开图形交互模式，从而反复操作fig_1
    plt.ion()

    global line_forecast
    fig_2.lines.remove(line_forecast[0])  # 删除之前的待预测指标图形

    global forecast_to_use
    global data_forecast_move
    date = forecast_to_use['日期']  # x轴，日期
    data_forecast_move = data_forecast_move+0.02  # y轴

    line_forecast = fig_2.plot(date, data_forecast_move, color='red')  # 绘图

    # 关闭图形交互模式
    plt.ioff()

    global canvas
    canvas_2.draw()

def move_down():

    global forecast_move_num
    forecast_move_num -= 1

    # 打开图形交互模式，从而反复操作fig_1
    plt.ion()

    global line_forecast
    fig_2.lines.remove(line_forecast[0])  # 删除之前的待预测指标图形

    global forecast_to_use
    global data_forecast_move
    date = forecast_to_use['日期']  # x轴，日期
    data_forecast_move = data_forecast_move-0.02  # y轴

    line_forecast = fig_2.plot(date, data_forecast_move, color='red')  # 绘图

    # 关闭图形交互模式
    plt.ioff()

    global canvas
    canvas_2.draw()

def selection_back_test():

    global listbox_forecast_var
    global name_forecast
    name_forecast = listbox_forecast_var.get(listbox_forecast_var.curselection())  # 获取Listbox控件中光标选择的值

    # 打开Listbox选中名字对应的数据，DataFrame，第一列是日期列
    global forecast_to_use
    forecast_to_use = forecast_result.parse(sheet_name=name_forecast)

    # 操作fig_2，调出待指标（name_forecast）对应的历史数据和预测数据，并绘制在图上
    global fig_2
    global line_history  # 历史数据用这个变量
    # 预测和回测均用这个变量
    global line_forecast  # 这里用于回测

    # 打开图形交互模式，从而反复操作fig_2
    plt.ion()

    # 判断line_history之前是否绘制，若存在，则删除之前的并绘制新的：若不存在，则绘制
    try:
        line_history
    except NameError:
        pass
    else:
        fig_2.lines.remove(line_history[0])  # 删除之前的待预测指标图形

    date = forecast_to_use['日期']  # x轴，日期
    data_history = forecast_to_use.iloc[:,1]  # y轴，历史数据

    line_history = fig_2.plot(date, data_history, color='blue')  # 绘图

    # 判断line_forecast之前是否绘制，若存在，则删除之前的并绘制新的：若不存在，则绘制
    try:
        line_forecast
    except NameError:
        pass
    else:
        fig_2.lines.remove(line_forecast[0])  # 删除之前的待预测指标图形

    date = forecast_to_use['日期']  # x轴，日期
    data_forecast = forecast_to_use.iloc[:,3]  # y轴，历史数据

    line_forecast = fig_2.plot(date, data_forecast, color='red')  # 绘图

    # 自动调整y坐标轴最大值和最小值，最大值等于line_history和line_forecast的最大值，最小值同理
    data_max = max(max(data_history), max(data_forecast))
    data_min = min(min(data_history), min(data_forecast))

    plt.ylim(data_min, data_max)

    # 关闭图形交互模式
    plt.ioff()

    # 回测向后移动数，后续可与日期关联
    global back_test_move_num
    back_test_move_num = 0

    # 总共回测期数
    global total_back_test_num
    total_back_test_num = forecast_to_use.shape[1]-3  # 前三列是日期，原始数据，预测数据，因此需要减去

    global canvas_2
    canvas_2.draw()

def previous_term():

    global forecast_to_use

    global back_test_move_num
    # 只有在back_test_move_num>0的时previous_term按钮才有效
    if back_test_move_num > 0:
        back_test_move_num -= 1

        # 打开图形交互模式，从而反复操作fig_1
        plt.ion()

        global line_forecast
        fig_2.lines.remove(line_forecast[0])  # 删除之前的待预测指标图形

        # 绘制新的line_forecast（回测）
        date = forecast_to_use['日期']  # x轴，日期
        data_back_test = forecast_to_use.iloc[:, back_test_move_num+3]  # y轴

        line_forecast = fig_2.plot(date, data_back_test, color='red')  # 绘图

        # 关闭图形交互模式
        plt.ioff()

        global canvas_2
        canvas_2.draw()

def next_term():

    global forecast_to_use
    # 得到回测总共的数量
    global total_back_test_num

    global back_test_move_num
    # 只有在back_test_move_num<total_back_test_num-1时（back_test_move_num是从0开始，所以要减1），next_term按钮才有效
    if back_test_move_num < total_back_test_num-1:

        back_test_move_num += 1

        # 打开图形交互模式，从而反复操作fig_1
        plt.ion()

        global line_forecast
        fig_2.lines.remove(line_forecast[0])  # 删除之前的待预测指标图形

        # 绘制新的line_forecast（回测）
        date = forecast_to_use['日期']  # x轴，日期
        data_back_test = forecast_to_use.iloc[:, back_test_move_num + 3]  # y轴

        line_forecast = fig_2.plot(date, data_back_test, color='red')  # 绘图

        # 关闭图形交互模式
        plt.ioff()

        global canvas_2
        canvas_2.draw()

# 历史回测自动播放
def auto_play():

    global back_test_move_num
    global total_back_test_num

    global pause_signal
    pause_signal = True  # 暂停信号

    # 自动终止播放条件，即回测移动数等于总的回测期数
    while back_test_move_num < total_back_test_num - 1:
        next_term()
        back_test_move_num += 1

        # 每次给出一期新回测后暂停一下
        time.sleep(0.2)

        if pause_signal == False:
            break

# tkinter主界面一直占用线程，需要开一个新线程实现auto_play功能
def auto_play_new_threading():

    th_1 = threading.Thread(target=auto_play)
    th_1.start()
    # th_1.join()

# 给auto_play()一个暂停信号
def pause():

    global pause_signal
    pause_signal = False

# GUI之2，展示result_3预测回测结果
def show_forecast():

    ## 创建窗体
    root_2 = tk.Tk()  # 创建窗体
    root_2.title('预测回测结果展示（TJD腾景数研）')  # 窗体名称
    root_2.geometry('1280x720')  # 窗体大小

    # Frame控件用于分割母窗体root_2
    frm_left = tk.Frame(root_2)
    frm_left.pack(side='left')
    frm_right = tk.Frame(root_2)
    frm_right.pack(side='right')

    frm_left_top = tk.Frame(frm_left)
    frm_left_top.pack(side='top')
    frm_left_bottom = tk.Frame(frm_left)
    frm_left_bottom.pack(side='bottom', fill='x', expand=1)

    frm_left_bottom_left = tk.Frame(frm_left_bottom)
    frm_left_bottom_left.pack(side='left', fill='x', expand=1)
    frm_left_bottom_right = tk.Frame(frm_left_bottom)
    frm_left_bottom_right.pack(side='right', fill='x', expand=1)

    # 按钮，用于选定和移动指标，并调出数据做图
    # 展示预测部分
    b1 = tk.Button(frm_left_bottom_left, text='预测结果', font=('宋体', 15),
                   width=10, height=2, command=selection_forecast)  # 注意！这里的按钮函数不能加（）
    b1.grid(row=1, column=2)

    b2 = tk.Button(frm_left_bottom_left, text='上移', font=('宋体', 15),
                   width=10, height=2, command=move_up)
    b2.grid(row=2, column=1)

    b3 = tk.Button(frm_left_bottom_left, text='下移', font=('宋体', 15),
                   width=10, height=2, command=move_down)
    b3.grid(row=2, column=3)

    # 展示回测部分
    b4 = tk.Button(frm_left_bottom_right, text='回测结果', font=('宋体', 15),
                   width=10, height=2, command=selection_back_test)
    b4.grid(row=1, column=6)

    b5 = tk.Button(frm_left_bottom_right, text='上一期', font=('宋体', 15),
                   width=10, height=2, command=previous_term)
    b5.grid(row=2, column=5)

    b6 = tk.Button(frm_left_bottom_right, text='下一期', font=('宋体', 15),
                   width=10, height=2, command=next_term)
    b6.grid(row=2, column=7)

    # 启动auto_play功能，但由于当前线程被tkinter主界面占用，需要通过auto_play_new_threading调用新线程
    b7 = tk.Button(frm_left_bottom_right, text='自动播放', font=('宋体', 15),
                   width=10, height=2, command=auto_play_new_threading)
    b7.grid(row=3, column=5)

    b8 = tk.Button(frm_left_bottom_right, text='暂停', font=('宋体', 15),
                   width=10, height=2, command=pause)
    b8.grid(row=3, column=7)

    # 滚动条，用于滚动列表框
    scorllbar = tk.Scrollbar(frm_right)
    scorllbar.pack(side='right', fill=tk.Y)

    # 列表框，存放预测回测的指标名称
    global forecast_result  # result_3的file变量
    forecast_result_sheet_names = forecast_result.sheet_names  # 给出result_3的sheet name,是一个List

    var_name = tk.StringVar()
    var_name.set(forecast_result_sheet_names)  # 指标名称传入var_name
    global listbox_forecast_var
    listbox_forecast_var = tk.Listbox(frm_right, listvariable=var_name, font=('Arial', 15),
                          yscrollcommand=scorllbar.set, width=40, height=30)  # 领先指标列表，var_lead_num移动这个里面的变量
    listbox_forecast_var.pack(fill=tk.BOTH)

    # 关联listbox的y方向view，使滚动条生效
    scorllbar.config(command=listbox_forecast_var.yview)

    # 画布，并将matplotlib绘制的图存放进来
    f_2 = plt.figure(figsize=(10, 5))  # 创建绘制图像
    global fig_2
    fig_2 = plt.subplot(1, 1, 1)

    global canvas_2
    canvas_2 = FigureCanvasTkAgg(f_2, frm_left_top)  # 把绘制图像显示到tkinter窗口上
    canvas_2.draw()  # canvas上显示，代替plt.show()
    canvas_2.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)  # 设置canvas位置

    ## 进入消息循环
    root_2.mainloop()

# 功能3：预测回测
def execute_function_3(show_data, parameter_sheet):

    # 传入参数信息，excel的行从-1开始数，列从0开始数
    select_fun_num = parameter_sheet.iloc[4, 4]  # 回归方程数选择

    forecast_switch = parameter_sheet.iloc[6, 4]  # 预测Switch
    forecast_sample_num = parameter_sheet.iloc[7, 4]  # 预测样本数（最近n个月，0为全样本）

    back_test_switch = parameter_sheet.iloc[9, 4]  # 回测Switch
    back_test_sample_num = parameter_sheet.iloc[10, 4]  # 回测样本数（n个月）

    # 创建保存结果的excel
    writer = pd.ExcelWriter('result_3.xlsx')  # 用pd.ExcelWriter的好处是，可以将多个DataFrame同时写入不同的Sheet中，而不会被覆盖
    result_3 = pd.DataFrame()  # 生成用于存放最终结果的DataFrame

    # 存放因变量名称，用于写入result_3的sheet name
    dependent_var_name = list()

    # 根据回归方程数分别拟合
    for i in range(select_fun_num):

        # 提取待拟合变量名称生成List
        var_name = parameter_sheet.iloc[15+6*i:19+6*i, 4]  ## 变量名参数传入
        var_name = var_name.dropna(axis=0, how='any')  # 删除没有填入变量的NaN
        var_name = list(var_name)

        # 将因变量名称传入dependent_var_name的List
        dependent_var_name.append(var_name[0])

        # 提取待拟合自变量滞后期生成List
        independent_var_num = len(var_name) - 1  # 自变量个数
        lead_num_list = list(parameter_sheet.iloc[16+6*i:16+6*i+independent_var_num, 5])  ## 变量领先期参数传入

        # 提取待预测数据并生成DataFrame
        data_to_use = show_data.loc[:, var_name]  # iloc按数字对DataFrame索引，loc按列名对DataFrame进行索引

        if forecast_switch == 1:
            print('步骤1：执行预测')
            # 预测
            result_3_forecast = forecast(data_to_use, lead_num_list, forecast_sample_num)
            if back_test_switch == 1:
                print('步骤2：执行回测')
                # 回测
                result_3_back_test = back_test(data_to_use, lead_num_list, back_test_sample_num)
                # 合并因变量、预测结果、回测结果
                result_3 = pd.concat([data_to_use.iloc[:, 0], result_3_forecast, result_3_back_test],
                                     axis=1, join_axes=[data_to_use.index])
                print('成功输出拟合方程'+str(i+1)+'的预测和回测结果')
            else:
                # 合并因变量、预测结果
                result_3 = pd.concat([data_to_use.iloc[:, 0], result_3_forecast],
                                     axis=1, join_axes=[data_to_use.index])
                print('成功输出拟合方程'+str(i+1)+'的预测结果')
        elif forecast_switch == 0:
            if back_test_switch == 1:
                print('步骤1：执行回测')
                # 回测
                result_3_back_test = back_test(data_to_use, lead_num_list, back_test_sample_num)
                # 合并因变量、回测结果
                result_3 = pd.concat([data_to_use.iloc[:, 0], result_3_back_test],
                                     axis=1, join_axes=[data_to_use.index])
                print('成功输出拟合方程'+str(i+1)+'的回测结果')
            else:
                print('请检查功能3的预测回测Switch设置')

        # 写入结果
        result_3.to_excel(writer, sheet_name='方程'+str(i + 1)+'_'+dependent_var_name[i])  # +str(dependent_var_name[i])   若写入多个sheet时，可使用sheet_name="a1"的参数
    # 保存结果
    writer.save()

    # 图形展示开关参数传入
    show_switch = parameter_sheet.iloc[12, 4]
    if show_switch == 1:
        global forecast_result
        forecast_result = pd.ExcelFile(r'result_3.xlsx')

        # 调用预测回测模块
        show_forecast()

###############################################################主函数##############################################################
def main():

    # 读取原始数据,要求原始数据是齐的
    parameter_sheet = pd.read_excel('Operation_Panel_beta3.0.xlsx', sheet_name='参数设置')  # 读取当前路径下Data，获得参数信息

    # 执行
    select_function = parameter_sheet.iloc[0, 1]  # 功能选择

    if select_function == 1:
        # 读取并执行功能1：数据处理
        data_sheet = pd.read_excel('Operation_Panel_beta3.0.xlsx', sheet_name='原始数据',
                                   index_col=0)  # 读取当前路径下Data，获得数据表，DataFrame格式
        result_1 = execute_function_1(data_sheet, parameter_sheet)

        # 输出结果
        writer = pd.ExcelWriter('result_1.xlsx')  # 用pd.ExcelWriter的好处是，可以将多个DataFrame同时写入不同的Sheet中，而不会被覆盖
        result_1.to_excel(writer)  # 若写入多个sheet时，可使用sheet_name="a1"的参数
        writer.save()

    elif select_function == 2:
        global data_show
        data_show = pd.read_excel('Operation_Panel_beta3.0.xlsx', sheet_name='待预测数据',
                                  header=0)  # 通过header=0参数，日期列保留，并将第一行作为列名
        execute_function_2()

    elif select_function == 3:
        show_data = pd.read_excel('Operation_Panel_beta3.0.xlsx', sheet_name='待预测数据',
                                  index_col=0)  # index_col=0，将第一行和第一列分别作为列index和行index
        execute_function_3(show_data, parameter_sheet)

    else:
        print('错误！请检查功能选择')

if __name__ == "__main__":
    main()
