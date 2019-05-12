import pandas_datareader.data as web
import datetime
import fix_yahoo_finance as yf

import pandas as pd
import xlwt
import re
import math

import numpy as np
import matplotlib.pyplot as plt
from matplotlib.dates import DateFormatter, WeekdayLocator, DayLocator, MONDAY, date2num
from mpl_finance import candlestick_ohlc

yf.pdr_override()

# 放的是 name 与 symbol 的对应文件，请用"|||"隔开
name_symbol_path = "name_symbol_dict"
# 输出数据文件夹
outfile_path = "output"

# 数据栏 | 可根据实际情况修改
COLUMNS = ["Open", "High", "Low", "Close", "Adj Close", "Volume"]
START_TIME = datetime.datetime(2016, 1, 1)
END_TIME = datetime.datetime.today()

# 不同指数按照不同等级标准划分 | 根据情况修改
class_diff_0_5 = ["JPY=X", "EURUSD=X", "CNY=X", "HKD=X",
                  "USDT-USD",
                  "^IRX", "^FVX", "^TNX", "^TYX", "^IXIC", "^DJI", "^GSPC", "^RUT", "^VIX", "IMOEX.ME", "^GDAXI", "^N225", "000001.SS", "^FTSE", "^HSI"]
class_diff_1_0 = ["GC=F", "SI=F", "PL=F", "HG=F", "PA=F", "CL=F"]
class_diff_3_0 = ["BTC-USD", "ETH-USD", "EOS-USD"]

# 分级涂色设置
num_color_dict = {-3: ('dark_green', 1), -2: ('sea_green', 1), -1: ('sea_green', 2),
                  0: ('gray25', 1),
                  1: ('coral', 2), 2: ('coral', 1), 3: ('dark_red', 1), float('nan'): ('white', 1)}


def read_name_symbole_file(path):
    ###
    # 读取 name 与 yahoo财经上的symbol对应文件
    name_symbol_tuple_list = list()
    with open(path, "r", encoding="utf-8") as nsdf:
        for line in nsdf:
            if line.startswith("#") == False and len(line.strip()) > 0:
                name, symbol = line.strip().split("|||")
                name_symbol_tuple_list.append((name, symbol))
    return name_symbol_tuple_list

def read_data_and_calculate_change_ratio(name_symbol_tuple_list):
    start = START_TIME
    end = END_TIME
    key = COLUMNS[4]
    full_data = dict()
    for name, symbol in name_symbol_tuple_list:
        data = web.get_data_yahoo(symbol, start, end)
        full_data[symbol] = data[key]
        print(name + ":" + str(len(data)))
    df = pd.DataFrame(full_data)
    df.plot(grid=True)
    df_path = "output\\{}.csv".format(key)
    df.to_csv(df_path)

    # df.fillna(method="ffill")
    columns = df.columns
    col_cnt = columns.size
    for i in range(0, columns.size):
        data = df.iloc[:, i]
        name = data.name
        df[name + '_change'] = data
        for j in range(1, data.size):
            if data[j] != float("nan"):
                previous = data[j]
                try:
                    for k in range(1, min(7, j + 1)):
                        if data[j - k] > 0:
                            previous = data[j - k]
                            break
                    if previous > 0 and data[j] > 0:
                        change_ratio = (data[j] - previous) / previous * 100
                        df.iloc[j, i + col_cnt] = change_ratio
                    else:
                        df.iloc[j, i + col_cnt] = 0
                except:
                    print("wrong:{}".format(data[j]))
            else:
                continue
        df[name] = df[name].map(lambda x: "{:.2f}".format(x) if x!=float("nan") else "")
        df[name + "_change"] = df[name +"_change"].map(lambda x: "({:+.2f}%)".format(x) if x!=float("nan") else "")
        df[name] = df[name].str.cat(df[name + "_change"])

    for column in columns:
        df.drop(column + "_change", axis=1, inplace=True)

    df_cal_path = "output\\{}_change.csv".format(key)
    df.to_csv(df_cal_path)
    # stock_change.head()
    # stock_change.plot(grid=True).axhline(y=0, color="black", lw=2)
    #
    # pandas_candlestick_ohlc(apple)
    # apple["20d"] = np.round(apple["Close"].rolling(window=20, center=False).mean(), 2)
    # pandas_candlestick_ohlc(apple.loc['2017-01-01':'2017-08-07', :], otherseries="20d")

    return df_cal_path

# 不需要的函数
# def pandas_candlestick_ohlc(dat, stick="day", otherseries=None):
#     """
#     :param dat: pandas DataFrame object with datetime64 index, and float columns "Open", "High", "Low", and "Close", likely created via DataReader from "yahoo"
#     :param stick: A string or number indicating the period of time covered by a single candlestick. Valid string inputs include "day", "week", "month", and "year", ("day" default), and any numeric input indicates the number of trading days included in a period
#     :param otherseries: An iterable that will be coerced into a list, containing the columns of dat that hold other series to be plotted as lines
#
#     This will show a Japanese candlestick plot for stock data stored in dat, also plotting other series if passed.
#     """
#     mondays = WeekdayLocator(MONDAY)  # major ticks on the mondays
#     alldays = DayLocator()  # minor ticks on the days
#     dayFormatter = DateFormatter('%d')  # e.g., 12
#
#     # Create a new DataFrame which includes OHLC data for each period specified by stick input
#     transdat = dat.loc[:, ["Open", "High", "Low", "Close"]]
#     if (type(stick) == str):
#         if stick == "day":
#             plotdat = transdat
#             stick = 1  # Used for plotting
#         elif stick in ["week", "month", "year"]:
#             if stick == "week":
#                 transdat["week"] = pd.to_datetime(transdat.index).map(lambda x: x.isocalendar()[1])  # Identify weeks
#             elif stick == "month":
#                 transdat["month"] = pd.to_datetime(transdat.index).map(lambda x: x.month)  # Identify months
#             transdat["year"] = pd.to_datetime(transdat.index).map(lambda x: x.isocalendar()[0])  # Identify years
#             grouped = transdat.groupby(list(set(["year", stick])))  # Group by year and other appropriate variable
#             plotdat = pd.DataFrame({"Open": [], "High": [], "Low": [],
#                                     "Close": []})  # Create empty data frame containing what will be plotted
#             for name, group in grouped:
#                 plotdat = plotdat.append(pd.DataFrame({"Open": group.iloc[0, 0],
#                                                        "High": max(group.High),
#                                                        "Low": min(group.Low),
#                                                        "Close": group.iloc[-1, 3]},
#                                                       index=[group.index[0]]))
#             if stick == "week":
#                 stick = 5
#             elif stick == "month":
#                 stick = 30
#             elif stick == "year":
#                 stick = 365
#
#     elif (type(stick) == int and stick >= 1):
#         transdat["stick"] = [np.floor(i / stick) for i in range(len(transdat.index))]
#         grouped = transdat.groupby("stick")
#         plotdat = pd.DataFrame(
#             {"Open": [], "High": [], "Low": [], "Close": []})  # Create empty data frame containing what will be plotted
#         for name, group in grouped:
#             plotdat = plotdat.append(pd.DataFrame({"Open": group.iloc[0, 0],
#                                                    "High": max(group.High),
#                                                    "Low": min(group.Low),
#                                                    "Close": group.iloc[-1, 3]},
#                                                   index=[group.index[0]]))
#
#     else:
#         raise ValueError(
#             'Valid inputs to argument "stick" include the strings "day", "week", "month", "year", or a positive integer')
#
#     # Set plot parameters, including the axis object ax used for plotting
#     fig, ax = plt.subplots()
#     fig.figsize = (15, 9)
#     fig.subplots_adjust(bottom=0.2)
#     if plotdat.index[-1] - plotdat.index[0] < pd.Timedelta('730 days'):
#         weekFormatter = DateFormatter('%b %d')  # e.g., Jan 12
#         ax.xaxis.set_major_locator(mondays)
#         ax.xaxis.set_minor_locator(alldays)
#     else:
#         weekFormatter = DateFormatter('%Y-%m-%d')
#     ax.xaxis.set_major_formatter(weekFormatter)
#
#     ax.grid(True)
#
#     # Create the candelstick chart
#     candlestick_ohlc(ax, list(
#         zip(list(date2num(plotdat.index.tolist())), plotdat["Open"].tolist(), plotdat["High"].tolist(),
#             plotdat["Low"].tolist(), plotdat["Close"].tolist())),
#                      colorup="red", colordown="green", width=stick * .4)
#
#     # Plot other series (such as moving averages) as lines
#     if otherseries != None:
#         if type(otherseries) != list:
#             otherseries = [otherseries]
#         dat.loc[:, otherseries].plot(ax=ax, lw=1.3, grid=True)
#
#     ax.xaxis_date()
#     ax.autoscale_view()
#     plt.setp(plt.gca().get_xticklabels(), rotation=45, horizontalalignment='right')
#
#     plt.show()

def set_style(num):
    try:
        color = num_color_dict[num][0]
        opa = num_color_dict[num][1]
        style = xlwt.XFStyle()
        # 初始化样式
        pattern = xlwt.Pattern()
        pattern.pattern = opa  # 设置底纹的图案索引，1为实心，2为50%灰色，对应为excel文件单元格格式中填充中的图案样式
        pattern.pattern_fore_colour = xlwt.Style.colour_map[color]  # 设置底纹的前景色，对应为excel文件单元格格式中填充中的背景色
        style.pattern = pattern  # 为样式设置图案
    except:
        print(num)
    return style

def color(path):

    df = pd.read_csv(path, header=0)
    output_fn = path.replace(".csv", "") + "_color.xls"

    # 新建一个excel文件
    file = xlwt.Workbook()
    # 新建一个sheet
    table = file.add_sheet('color', cell_overwrite_ok=True)
    columns = df.columns
    col_count = columns.size

    # 按照分级确定涂色
    num_style_dict = dict()
    for key in num_color_dict.keys():
        num_style_dict[key] = set_style(key)

    style1 = xlwt.XFStyle()

    font = xlwt.Font()
    font.bold = True

    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER  # 水平方向
    alignment.vert = xlwt.Alignment.VERT_TOP  # 垂直方向

    style1.alignment = alignment
    style1.font = font

    index = df.iloc[:, 0]
    for i in range(col_count):
        table.write(0, i, columns[i], style1)
    for i in range(index.size):
        table.write(i+1, 0, index[i], style1)

    for j in range(1, col_count):
        data = df.iloc[:, j]
        name = data.name
        num = 0
        for i in range(data.size):
            try:
                ratio = re.findall(r'[(](.+)[)]', data[i])[0].replace('%', '')
                float_ratio = float(ratio)
                if name in class_diff_0_5:
                    if float_ratio >= 0:
                        num = math.ceil(float_ratio / 0.5)
                    elif float_ratio < 0:
                        num = math.floor(float_ratio / 0.5)
                elif name in class_diff_1_0:
                    if float_ratio >= 0:
                        num = math.ceil(float_ratio / 1.0)
                    elif float_ratio < 0:
                        num = math.floor(float_ratio / 1.0)
                elif name in class_diff_3_0:
                    if float_ratio >= 0:
                        num = math.ceil(float_ratio / 3.0)
                    elif float_ratio < 0:
                        num = math.floor(float_ratio / 3.0)
                if num > 3 and num != float('nan'):
                    num = 3
                elif num < -3 and num != float('nan'):
                    num = -3
                style = num_style_dict[num]
                table.write(i+1, j, data[i], style)  # 使用样式
            except:
                print('{0} - i: {1}, j: {2}'.format(data[i], i, j))
    table.set_panes_frozen(True)
    table.set_horz_split_pos(1)
    table.set_vert_split_pos(1)
    table.col(0).width = 5000
    file.save(output_fn)

if __name__ == "__main__":
    name_symbol_tuple_list = read_name_symbole_file(name_symbol_path)
    df_cal_path = read_data_and_calculate_change_ratio(name_symbol_tuple_list)
    color(df_cal_path)
    # color("Adj Close_change.csv")