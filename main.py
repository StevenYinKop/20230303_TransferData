import os

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

output_filename = 'total.xlsx'

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
script_path = os.path.abspath(__file__)
base_dir = os.path.dirname(script_path)


def assemble_data_list_item(field, data, index_dict):
    return [
        index_dict[title_indicator_name],  # 数据指标名称
        str(field),  # 字段名称
        int(index_dict[title_year]),  # 数据年度
        data  # 数据
    ]


# 校验索引文件的格式, 防止索引文件中，输入的数据质量差，使程序崩溃
def validate_index_file_parameters(headers, index_filename):
    for title_str in title_list:
        if title_str not in headers:
            raise ValueError(f'索引文件"{index_filename}"中缺失必要参数："{title_str}"，请检查。')


# 定义一个解析数据文件的函数
def parse_data_file(file_path, index_row_dict):
    sub_data_list = []
    print(f'开始处理文件{file_path}...')
    _, extension = os.path.splitext(file_path)
    if extension == '.xlsx':
        print('文件是xlsx格式，使用新excel的解析方式...请稍后...')
        workbook = load_workbook(file_path)
        sheet = workbook[index_row_dict[title_sheet_name]] \
            if index_row_dict[title_sheet_name] not in [None, ''] \
            else workbook.worksheets[0]
        print(f'文件加载成功，当前操作的sheet页为{sheet.title}')
        if index_row_dict[title_is_horizontal_or_vertical] == '横表':
            print(f'以"{index_row_dict[title_is_horizontal_or_vertical]}"'
                  f'的方式解析当前sheet页的第{int(index_row_dict[title_start_at])}'
                  f'列至第{int(index_row_dict[title_end_at])}列...')

            for col in sheet.iter_cols(min_col=max(int(index_row_dict[title_start_at]), sheet.min_column),
                                       max_col=min(int(index_row_dict[title_end_at]), sheet.max_column)):
                print(f'开始解析第{col[0].column}列...')
                item = assemble_data_list_item(col[int(index_row_dict[title_field_index]) - 1].value,
                                               col[int(index_row_dict[title_data_index]) - 1].value,
                                               index_row_dict)
                sub_data_list.append(item)
                print(f'第{col[0].column}列解析完毕！, 数据为：{item}')
        elif index_row_dict[title_is_horizontal_or_vertical] == '竖表':
            print(f'以"{index_row_dict[title_is_horizontal_or_vertical]}"'
                  f'的方式解析当前sheet页的第{int(index_row_dict[title_start_at])}'
                  f'行至第{int(index_row_dict[title_end_at])}行...')
            # todo 待测试
            for row in sheet.iter_rows(min_col=max(int(index_row_dict[title_start_at]), sheet.min_row),
                                       max_col=min(int(index_row_dict[title_end_at]), sheet.max_row)):
                print(f'开始解析第{row[0].column}行...')
                item = assemble_data_list_item(row[int(index_row_dict[title_field_index]) - 1].value,
                                               row[int(index_row_dict[title_data_index]) - 1].value,
                                               index_row_dict)
                sub_data_list.append(item)
                print(f'第{row[0].column}行解析完毕！, 数据为：{item}')

    elif extension == '.xls':
        print('文件是xls格式，使用旧excel的解析方式...请稍后...')
        workbook = open_workbook(file_path)
        sheet = workbook.sheet_by_name(index_row_dict[title_sheet_name]) \
            if index_row_dict[title_sheet_name] not in [None, ''] \
            else workbook.sheet_by_index(0)
        print(f'文件加载成功，当前操作的sheet页为{sheet.name}')

        if index_row_dict[title_is_horizontal_or_vertical] == '横表':
            print(f'以"{index_row_dict[title_is_horizontal_or_vertical]}"'
                  f'的方式解析当前sheet页的第{int(index_row_dict[title_start_at])}'
                  f'列至第{int(index_row_dict[title_end_at])}列...')
            for col in range(int(index_row_dict[title_start_at]) - 1, int(index_row_dict[title_end_at])):
                print(f'开始解析第{col + 1}列...')
                item = assemble_data_list_item(sheet.cell(int(index_row_dict[title_field_index]) - 1, col).value,
                                               sheet.cell(int(index_row_dict[title_data_index]) - 1, col).value,
                                               index_row_dict)
                sub_data_list.append(item)
                print(f'第{col + 1}列解析完毕！, 数据为：{item}')
        elif index_row_dict[title_is_horizontal_or_vertical] == '竖表':
            print(f'以"{index_row_dict[title_is_horizontal_or_vertical]}"'
                  f'的方式解析当前sheet页的第{int(index_row_dict[title_start_at])}'
                  f'行至第{int(index_row_dict[title_end_at])}行...')
            for row in range(int(index_row_dict[title_start_at]) - 1, int(index_row_dict[title_end_at])):
                print(f'开始解析第{row + 1}行...')
                item = assemble_data_list_item(sheet.cell(row, int(index_row_dict[title_field_index]) - 1).value,
                                               sheet.cell(row, int(index_row_dict[title_data_index]) - 1).value,
                                               index_row_dict)
                sub_data_list.append(item)
                print(f'第{row + 1}行解析完毕！, 数据为：{item}')
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
        headers = [cell.value for cell in sheet[0]]
        validate_index_file_parameters(headers, index_filename)
        for row in sheet.iter_rows(min_row=2, values_only=True):
            index_row_dict = {headers[i]: row[i] for i in range(len(headers))}
            if title_file_path not in index_row_dict:
                raise ValueError(f'{index_filename}中不存在名为{title_file_path}的列，无法找到对应的数据文件！')
            data_list.extend(parse_data_file(f'{base_dir}/{index_row_dict[title_file_path]}', index_row_dict))

    # 如果是xls后缀名的文件，走以下逻辑
    elif extension == '.xls':
        workbook = open_workbook(index_filename)
        sheet = workbook.sheet_by_index(0)
        headers = [cell.value for cell in sheet[0]]
        validate_index_file_parameters(headers, index_filename)
        for row in range(1, sheet.nrows):
            index_row_dict = {headers[i]: sheet.cell_value(row, i) for i in range(len(headers))}
            if title_file_path not in index_row_dict:
                raise ValueError(f'{index_filename}中不存在名为{title_file_path}的列，无法找到对应的数据文件！')
            data_list.extend(parse_data_file(f'{base_dir}/{index_row_dict[title_file_path]}', index_row_dict))


    # 如果既不是xls后缀名，也不是xlsx后缀名，说明输入的索引文件有问题。
    else:
        print("输入的文件不是一个标准的excel，请使用.xls或者.xlsx文件")


def write_to_output_file():
    workbook_total = Workbook()
    sheet_total = workbook_total.active
    for id_r, item in enumerate(data_list):
        for id_c, value in enumerate(item):
            sheet_total.cell(id_r + 1, id_c + 1, value)
    workbook_total.save(f'{base_dir}/total.xlsx')


# 定义python的入口文件，代码从此处开始运行⬇️
def main():
    # 获得索引文件
    index_filename = 'index.xls'

    # 解析索引文件
    parse_index_file(index_filename)
    print(data_list)
    write_to_output_file()


if __name__ == "__main__":
    main()
