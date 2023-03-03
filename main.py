import os
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from xlrd import open_workbook

# 定义标题名称
title_author = '数据采集人'
title_file_path = '数据文件名称'
title_sheet_name = 'Sheet页名称(默认为第一页）'
title_indicator_name = '采集指标名称'
title_year = '采集年度'
title_start_at = '采集起始行'
title_end_at = '采集终止行'
title_field_index = '字段名称列号'
title_data_index = '数据列号'
title_is_horizontal_or_vertical = '横表/竖表'

# 声明一个结果集，后续的结果都会存在这个列表中
data_list = []
script_path = os.path.abspath(__file__)
base_dir = os.path.dirname(script_path)

#定义一个解析数据文件的函数
def parse_data_file(file_path, row_dict):
    _, extension = os.path.splitext(file_path)
    if extension == '.xlsx':
        pass
    elif extension == '.xls':
        workbook = open_workbook(file_path)
        sheet = workbook.sheet_by_name(row_dict[title_sheet_name]) if row_dict[title_sheet_name] not in [None, ''] else workbook.sheet_by_name(0)

        headers = [cell.value for cell in sheet[0]]
        for row in range(1, sheet.nrows):
            row_dict = {headers[i]: sheet.cell_value(row, i) for i in range(len(headers))}


    if row_dict[title_is_horizontal_or_vertical == '横表']:
        pass
    # todo
    elif row_dict[title_is_horizontal_or_vertical == '竖表']:
        pass
    # todo
    else:
        print("输入的文件不是一个标准的excel，请使用.xls或者.xlsx文件")

# 定义一个解析索引文件的函数
def parse_index_file(index_filename):
    _, extension = os.path.splitext(index_filename)

    # 如果是xlsx后缀名的文件，走这个if逻辑
    if extension == '.xlsx':
        workbook = load_workbook(index_filename)
        sheet = workbook.active
        headers = [cell.value for cell in sheet[0]]
        for row in sheet.iter_rows(min_row=2, values_only=True):
            row_dict = {headers[i]: row[i] for i in range(len(headers))}
            if title_file_path not in row_dict:
                raise ValueError(f'{index_filename}中不存在名为{title_file_path}的列，无法找到对应的数据文件！')
            data_list.extend(parse_data_file(f'{base_dir}/{row_dict[title_file_path]}', row_dict))

    # 如果是xls后缀名的文件，走以下逻辑
    elif extension == '.xls':
        workbook = open_workbook(index_filename)
        sheet = workbook.sheet_by_index(0)
        headers = [cell.value for cell in sheet[0]]
        for row in range(1, sheet.nrows):
            row_dict = {headers[i]: sheet.cell_value(row, i) for i in range(len(headers))}
            if title_file_path not in row_dict:
                raise ValueError(f'{index_filename}中不存在名为{title_file_path}的列，无法找到对应的数据文件！')
            data_list.extend(parse_data_file(f'{base_dir}/{row_dict[title_file_path]}', row_dict))


    # 如果既不是xls后缀名，也不是xlsx后缀名，说明输入的索引文件有问题。
    else:
        print("输入的文件不是一个标准的excel，请使用.xls或者.xlsx文件")


# 定义python的入口文件，代码从此处开始运行⬇️
def main():
    # 获得索引文件
    index_filename = 'index.xls'

    # 解析索引文件
    parse_index_file(index_filename)

if __name__ == "__main__":
    main()

