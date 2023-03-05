import os
import sys

from openpyxl import load_workbook, Workbook
from xlrd import open_workbook

# 定义标题名称
title_author = '数据采集人'
title_file_path = '数据文件名称'
title_sheet_name = 'Sheet页名称(默认为第一页）'
title_indicator_name = '采集指标名称'
title_year = '数据年度'
title_start_at = '采集起始行/列'
title_end_at = '采集终止行/列'
title_field_index = '字段名称列/行号'
title_data_index = '数据列/行号'
title_is_horizontal_or_vertical = '横表/竖表'

script_path = os.path.abspath(__file__)
base_dir = f'{os.path.dirname(script_path)}{os.sep}工作目录'

output_filename = f'{base_dir}{os.sep}total.xlsx'
error_filename = f'{base_dir}{os.sep}错误日志.txt'
error_list = []
error_index_count = 0

title_list = [
    title_author,
    title_file_path,
    title_sheet_name,
    title_indicator_name,
    title_year,
    title_start_at,
    title_end_at,
    title_field_index,
    title_data_index,
    title_is_horizontal_or_vertical
]

# 声明一个结果集，后续的结果都会存在这个列表中
data_list = [
    ['数据指标名称',
     '字段名称',
     '数据年度',
     '数据']
]


def get_format_error_cell(error_str, row_number):
    return f'[-----第{row_number}行-----]：{error_str}'


def output_error_list():
    if len(error_list) == 0:
        return
    error_file = open(error_filename, 'w')
    for error_item in error_list:
        error_file.write(f'{error_item}\n')
    print(f"程序执行过程中发现错误，请查看{error_filename}")
    error_file.close()
    sys.exit()
    # exit()


def standardize_index(index_row_dict, get_max):
    return max(int(index_row_dict[title_start_at]), 1), \
           min(int(index_row_dict[title_end_at]), get_max())


def assemble_data_list_item(field, data, index_dict):
    return [
        index_dict[title_indicator_name],  # 数据指标名称
        str(field),  # 字段名称
        int(index_dict[title_year]),  # 数据年度
        data  # 数据
    ]


# 校验索引文件的格式, 防止索引文件中，输入的数据质量差，使程序崩溃
def validate_index_file_parameters(headers, index_filename):
    print(f'正在检查索引文件"{index_filename}"中的数据有效性...')
    for title_str in title_list:
        if title_str not in headers:
            error_list.append(f'索引文件"{index_filename}"中缺失必要参数列："{title_str}"，请检查。')
    output_error_list()


def validate_index_data(index_dict, row):
    error_sub_list = []
    for title_str in [title_year, title_start_at, title_end_at, title_data_index, title_field_index]:
        try:
            int(float((index_dict[title_str])))
        except ValueError:
            error_sub_list.append(get_format_error_cell(f' "{title_str}" 不是一个数字格式，无法解析！', row))
    if index_dict[title_is_horizontal_or_vertical] not in ['横表', '竖表']:
        error_sub_list.append(get_format_error_cell(f' "{title_is_horizontal_or_vertical}"中只能填"横表"或者"竖表"！', row))
    if not os.path.exists(f'{base_dir}{os.sep}{str(index_dict[title_file_path])}'):
        error_sub_list.append(get_format_error_cell(f' "{title_file_path}" 文件不存在，无法解析！', row))
    else:
        file_path = f'{base_dir}{os.sep}{str(index_dict[title_file_path])}'
        _, extension = os.path.splitext(file_path)
        if extension == '.xlsx':
            workbook = load_workbook(file_path)
            if index_dict[title_sheet_name] not in workbook.sheetnames and index_dict[title_sheet_name] is not None:
                error_sub_list.append(
                    get_format_error_cell(f'{title_file_path}文件中不存在名为{index_dict[title_sheet_name]}的Sheet页，无法解析！', row))
            workbook.close()
        elif extension == '.xls':
            workbook = open_workbook(file_path)
            if index_dict[title_sheet_name] not in workbook.sheet_names() and str(index_dict[title_sheet_name]) not in [
                'None', '']:
                error_sub_list.append(
                    get_format_error_cell(f'{title_file_path}文件中不存在名为{index_dict[title_sheet_name]}的Sheet页，无法解析！', row))
            workbook.release_resources()
    if len(error_sub_list) != 0:
        error_sub_list.append(get_format_error_cell(f'跳过第{row}行...\n', row))
    error_list.extend(error_sub_list)
    return len(error_sub_list) == 0


# 定义一个解析数据文件的函数
def parse_data_file(file_path, index_row_dict):
    sub_data_list = []
    print(f'开始处理文件{file_path}...')
    _, extension = os.path.splitext(file_path)
    field_index = int(index_row_dict[title_field_index])
    data_index = int(index_row_dict[title_data_index])
    if extension == '.xlsx':
        print('文件是xlsx格式，使用新excel的解析方式...请稍后...')
        workbook = load_workbook(file_path)
        sheet = workbook[index_row_dict[title_sheet_name]] \
            if index_row_dict[title_sheet_name] not in [None, ''] \
            else workbook.worksheets[0]
        print(f'文件加载成功，当前操作的sheet页为{sheet.title}')
        if index_row_dict[title_is_horizontal_or_vertical] == '横表':
            start_at, end_at = standardize_index(index_row_dict, lambda: sheet.max_column)
            print(f'以"{index_row_dict[title_is_horizontal_or_vertical]}"'
                  f'的方式解析当前sheet页的第{start_at}列至第{end_at}列...')
            for col in range(start_at, end_at + 1):
                print(f'开始解析第{col}列...')
                item = assemble_data_list_item(sheet.cell(field_index, col).value,
                                               sheet.cell(data_index, col).value,
                                               index_row_dict)
                sub_data_list.append(item)
                print(f'第{col}列解析完毕！, 数据为：{item}')
        elif index_row_dict[title_is_horizontal_or_vertical] == '竖表':
            start_at, end_at = standardize_index(index_row_dict, lambda: sheet.max_row)
            print(f'以"{index_row_dict[title_is_horizontal_or_vertical]}"'
                  f'的方式解析当前sheet页的第{start_at}行至第{end_at}行...')
            for row in range(start_at, end_at + 1):  # sheet.iter_rows(min_col=start_at, max_col=end_at):
                print(f'开始解析第{row}行...')
                # sheet[]
                item = assemble_data_list_item(sheet.cell(row, field_index).value,
                                               sheet.cell(row, data_index).value,
                                               index_row_dict)
                sub_data_list.append(item)
                print(f'第{row}行解析完毕！, 数据为：{item}')

    elif extension == '.xls':
        print('文件是xls格式，使用旧excel的解析方式...请稍后...')
        workbook = open_workbook(file_path)
        sheet = workbook.sheet_by_name(index_row_dict[title_sheet_name]) \
            if index_row_dict[title_sheet_name] not in [None, ''] \
            else workbook.sheet_by_index(0)
        print(f'文件加载成功，当前操作的sheet页为{sheet.name}')

        if index_row_dict[title_is_horizontal_or_vertical] == '横表':
            start_at, end_at = standardize_index(index_row_dict, lambda: sheet.ncols)
            print(f'以"{index_row_dict[title_is_horizontal_or_vertical]}"'
                  f'的方式解析当前sheet页的第{start_at}列至第{end_at}列...')
            for col in range(start_at, end_at + 1):
                print(f'开始解析第{col}列...')
                item = assemble_data_list_item(sheet.cell_value(field_index - 1, col - 1),
                                               sheet.cell_value(data_index - 1, col - 1),
                                               index_row_dict)
                sub_data_list.append(item)
                print(f'第{col}列解析完毕！, 数据为：{item}')
        elif index_row_dict[title_is_horizontal_or_vertical] == '竖表':
            start_at, end_at = standardize_index(index_row_dict, lambda: sheet.nrows)
            print(f'以"{index_row_dict[title_is_horizontal_or_vertical]}"'
                  f'的方式解析当前sheet页的第{start_at}行至第{end_at}行...')
            for row in range(start_at, end_at + 1):
                print(f'开始解析第{row}行...')
                item = assemble_data_list_item(sheet.cell_value(row - 1, field_index - 1),
                                               sheet.cell_value(row - 1, data_index - 1),
                                               index_row_dict)
                sub_data_list.append(item)
                print(f'第{row}行解析完毕！, 数据为：{item}')
    else:
        print(f'\n错误！！！ === 输入的文件" {file_path} "不是一个标准的excel，请使用.xls或者.xlsx文件 === ', end='\n\n')
    return sub_data_list


# 定义一个解析索引文件的函数
def parse_index_file(index_filename):
    _, extension = os.path.splitext(index_filename)

    # 如果是xlsx后缀名的文件，走这个if逻辑
    if extension == '.xlsx':
        workbook = load_workbook(index_filename)
        sheet = workbook.active
        headers = [cell.value for cell in sheet[1]]
        validate_index_file_parameters(headers, index_filename)
        for row in sheet.iter_rows(min_row=2):
            index_row_dict = {headers[i]: row[i].value for i in range(len(headers))}
            if validate_index_data(index_row_dict, row[0].row):
                data_list.extend(
                    parse_data_file(f'{base_dir}{os.sep}{index_row_dict[title_file_path]}', index_row_dict))

    # 如果是xls后缀名的文件，走以下逻辑
    elif extension == '.xls':
        workbook = open_workbook(index_filename)
        sheet = workbook.sheet_by_index(0)
        headers = [cell.value for cell in sheet[0]]
        validate_index_file_parameters(headers, index_filename)
        for row in range(1, sheet.nrows):
            index_row_dict = {headers[i]: sheet.cell_value(row, i) for i in range(len(headers))}
            if validate_index_data(index_row_dict, row):
                data_list.extend(
                    parse_data_file(f'{base_dir}{os.sep}{index_row_dict[title_file_path]}', index_row_dict))

    # 如果既不是xls后缀名，也不是xlsx后缀名，说明输入的索引文件有问题。
    else:
        print("输入的文件不是一个标准的excel，请使用.xls或者.xlsx文件")


def write_to_output_file():
    workbook_total = Workbook()
    sheet_total = workbook_total.active
    for id_r, item in enumerate(data_list):
        for id_c, value in enumerate(item):
            sheet_total.cell(id_r + 1, id_c + 1, value)
            print(value, end='\t')
        print()
    workbook_total.save(output_filename)
    if len(data_list) > 1:
        print(f'数据全部处理完毕，写入文件： {output_filename}')
        workbook_total.close()
    print(f'共写入数据{len(data_list) - 1}条！')


# 定义python的入口文件，代码从此处开始运行⬇️
def main():
    # 获得索引文件
    global base_dir
    global output_filename
    global error_filename
    default_index_filename = f'{base_dir}{os.sep}index.xls'
    print("欢迎使用数据迁移程序！")
    index_filename = input(f'请输入索引文件的路径，按回车键确认(默认路径为：{default_index_filename}):\n')
    if index_filename == '':
        index_filename = default_index_filename
    base_dir = f'{os.path.dirname(index_filename)}{os.sep}'
    output_filename = f'{base_dir}{os.sep}total.xlsx'
    error_filename = f'{base_dir}{os.sep}错误日志.txt'
    input(f'请将所有的数据文件，放置在"{base_dir}"文件夹下，按回车键确认\n')
    if os.path.exists(index_filename):
        # 解析索引文件
        parse_index_file(index_filename)
        write_to_output_file()
        output_error_list()
    else:
        print(f'找不到"{index_filename}"，请检查文件是否存在！')


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(e)
    finally:
        input("按下任意键结束程序...")
