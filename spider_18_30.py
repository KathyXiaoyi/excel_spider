# coding:utf-8
import re
import xlrd
from item import Item
from util import month_map, is_num, is_month, print_item


# 从表格的1-3行中匹配出month
def match_month(sheet_param):
    for index in range(0, 3):
        row_content = sheet_param.row_values(index)
        for content_data in row_content:
            if content_data:
                data_list = content_data.split(' ')
                return month_map(data_list[0])


# 从表格的5-7行中匹配出commodity
def match_commodity(sheet_param):
    for row_index in range(4, 7):
        row_content = sheet_param.row_values(row_index)
        for datum in row_content:
            if datum:
                contents = datum.split(u' ')
                if contents[1] == 'Metric' or contents[0] == 'WASDE':
                    break
                end = datum.find('Supply') - 1
                begin = datum.find(contents[1])
                commodity = datum[begin:end]
                return commodity


# 从5-9行中匹配出unit
def match_unit(sheet_param):
    for i in range(5, 9):
        data = sheet_param.row_values(i)
        for datum in data:
            if datum:
                match = re.match(r'\(.*\)', datum)
                if match:
                    return match.group()[1:-1]


# 在8-11行中匹配出market_year
def match_market_year(sheet_param):
    for i in range(8, 11):
        data = sheet_param.row_values(i)
        for datum in data:
            if datum:
                if_match = re.match(r'\d{4}\/\d{2}', datum)
                if if_match:
                    return if_match.group()


# 找到attribute所在的行数下标
def find_attribute_line(sheet_param):
    # 在表格的8-10中寻找attribute行
    for i in range(8, 11):
        data = sheet_param.row_values(i)
        for datum in data:
            if datum:
                # 匹配时间-时间所在行即为attribute所在行
                if_match = re.match(r'\d{4}\/\d{2}', datum)
                if if_match:
                    return i


# 尝试从某一行的0-7列中匹配出指定的month
def find_specific_month(sheet_param, row_index, col_begin_index, col_end_index):
    for i in range(col_begin_index, col_end_index):
        content = sheet_param.cell(row_index, i).__str__()
        if content[0:5] != 'empty':
            if is_month(content[7:-1]):
                return content[7:-1]
    return None


def match_data(sheet_param, calendar_year_param):
    list = []
    count = 0
    max_rows = sheet_param.nrows
    # 寻找attribute所在行的下标
    attribute_line = find_attribute_line(sheet_param)
    # 读取attribute
    data_attribute = sheet_param.row_values(attribute_line)
    # 匹配出Default month
    month = match_month(sheet_param)
    # 匹配出commodity
    commodity = match_commodity(sheet_param)
    # 匹配出unit
    unit = match_unit(sheet_param)
    # 匹配出market_year
    market_year = match_market_year(sheet_param)
    index_temp = 0
    for i in range(8, max_rows-1):
        # 匹配country
        data = sheet_param.row_values(i)
        data_up = sheet_param.row_values(i-1)
        j = 0
        if data[0] and not is_month(data[0]) and not is_num(data[0]):
            country = data[0]
            index_temp = 0
        elif data[1] and not is_month(data[1]) and not is_num(data[1]):
            country = data[1]
            index_temp = 1
        elif data[2] and not is_month(data[2]) and not is_num(data[2]):
            country = data[2]
            index_temp = 2
        elif data[3] and not is_month(data[3]) and not is_num(data[3]):
            country = data[3]
            index_temp = 3
        else:
            country = data_up[index_temp]
            # page 25中的特殊情况处理
            if country is None or country.strip() == '':
                country = sheet_param.row_values(i-2)[index_temp]
        # 从1-6列中尝试匹配指定的month
        max_col_index = sheet_param.ncols
        if max_col_index < 9:
            month_temp = find_specific_month(sheet_param, i, 0, max_col_index)
        else:
            month_temp = find_specific_month(sheet_param, i, 0, 9)
        if month_temp is not None:
            month = month_temp
        for datum in data:
            if datum or datum == 0.0:
                if is_num(datum):
                    item = Item()
                    item.Commodity = commodity
                    item.Country = country
                    item.MarketYear = market_year
                    item.CalendarYear = calendar_year_param
                    item.Month = month
                    # item.Attribute = data_attribute[j]
                    item.Attribute = sheet_param.cell(attribute_line, j).__str__()[7:-1]
                    item.Unit = unit
                    item.Value = str(datum)
                    list.append(item)
                else:
                    # 尝试寻找新的market_year
                    if_match = re.match(r'\d{4}\/\d{2}', datum)
                    if if_match:
                        market_year = datum
            j += 1
    return list


if __name__ == '__main__':
    # 打开excel
    file_path_param = 'D:/data/wasde-01-10-2014.xls'
    workbook = xlrd.open_workbook(file_path_param)
    sheet = workbook.sheet_by_name('Page 23')
    # print find_specific_month(sheet, 23, 0, 7)
    # print find_specific_month(sheet, 22, 0, 7)
    # print match_commodity(sheet)
    # print match_month(sheet)
    # print match_market_year(sheet)
    # print match_unit(sheet)
    # 读取calendar_year
    calendar_year = file_path_param[-8:-4]
    item_list = match_data(sheet, calendar_year)
    print_item(item_list)