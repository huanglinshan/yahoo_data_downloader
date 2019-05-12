## 文件说明
    主程序：                      extract_data.py
    指数和yahoo财经代码文件:       name_symbol_dict

## 使用说明
    1 可以选择以下列的内容输出
         COLUMNS = ["Open", "High", "Low", "Close", "Adj Close", "Volume"]
      分别为：开盘价，最高价，最低价，收盘价，调整后的收盘价，成交量 
    2 在
        def read_data_and_calculate_change_ratio(name_symbol_tuple_list):
            ...
            key = COLUMNS[4] 
      中设置需要下载数据的列名称。比如这里 COLUMNS[4] 就是指"Adj Close" 
    3 运行 extract_data.py，可以按照比率给出变化大小，最终文件名称为"XXX_change_color.csv"