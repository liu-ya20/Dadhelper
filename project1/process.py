import os
import pandas as pd
import re

# 设置原始数据文件夹路径
data_folder_to_process = '2024-02-02-火_达拉特热电厂#1-#4机组#3机组(出清结果)'
data_collection = 'D:/Pycoding/project1/data'
result_path = 'D:/Pycoding/project1/result/result.xlsx'


# 初始化汇总数据
summary_data = pd.DataFrame()
column_names = ['日期', '厂名', '机组', '上网电量', '电能电价', '电能电费', '实时出清电力', '实时节点电价', '中长期合约电量',
       '中长期合约电价', '省间日前出清电力', '省间日前出清电价', '省间日内出清电力', '省间日内出清电价', '启动补偿费用',
       '必开机组成本补偿费用', '顶峰能力考核费用', '顶峰能力考核返还']

extra_column_names = ['合约电价']

template_data = pd.DataFrame(columns = column_names)

# 循环处理初始数据文件
data_path = os.path.join(data_collection, data_folder_to_process)
data_files = [f for f in os.listdir(data_path) if f.endswith('.xlsx')]


# 定义提取信息的正则表达式
pattern = re.compile(r'(\d{4}-\d{2}-\d{2})-火_(\S+热电厂)#\S+机组(#\d+机组)\(出清结果\)')
selected_column_index = [97]
selected_row_index = list(range(1,16))
info_df = pd.DataFrame({'日期': [],'厂名':[], '机组':[]})
for data_file in data_files:
    # 分析数据名字
    match = pattern.match(data_file)
    if match:
        date = match.group(1)
        name = match.group(2)
        unit_number = match.group(3)
    else:
        print("文件名格式非常规，无法提取信息")

    df = pd.DataFrame({'日期':[date], '厂名':[name], '机组':[unit_number]})
    info_df = pd.concat([info_df, df])

    # 读取初始数据文件
    data_file_path = os.path.join(data_path, data_file)
    data = pd.read_excel(data_file_path)

    # 在这里添加代码以从初始数据中提取特定行，并将其添加到汇总数据中
    # 这里提取所有有效信息行的总计数据
    specific_data = data.iloc[selected_row_index, selected_column_index]

    # 将提取的数据添加到汇总数据中
    summary_data = pd.concat([summary_data, specific_data], axis = 1)

target_data = summary_data.transpose().reset_index(drop = True)
# new_columns = ['顶峰能力考核费用','顶峰能力考核返还']
# target_data.columns = column_names

target_info = info_df.reset_index(drop=True)
# 将信息与数据拼接
joined = target_info.join(target_data)
joined.columns = column_names

# 放入空模板中
joined[extra_column_names[0]] = joined['中长期合约电量']*joined['中长期合约电价']
joined.to_excel(result_path, index=False)
