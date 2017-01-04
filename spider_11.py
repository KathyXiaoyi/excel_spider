# coding:utf-8
from util import month_map, is_num, get_calendar_year, is_month, get_file_list
from item import Item
import re
import xlrd


# 从1-3行中匹配出Default month
def find_default_month(sheet_param):
    for index in range(0, 3):
        row_content = sheet_param.row_values(index)
        for content_data in row_content:
            if content_data:
                data_list = content_data.split(' ')
                return month_map(data_list[0])


# 从表格的4-7行中匹配出commodity
def find_commodity(sheet_param):
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


# 匹配出下半部分的commodity
def find_second_commodity(sheet_param, row_index):
    row_content = sheet_param.row_values(row_index)
    for datum in row_content:
        if datum:
            contents = datum.split(u' ')
            end = datum.find('by') - 1
            begin = datum.find(contents[1])
            commodity = datum[begin:end]
            return commodity


# 匹配出下半部分的commodity
def find_second_country(sheet_param, row_index):
    row_content = sheet_param.row_values(row_index)
    for datum in row_content:
        if datum:
            contents = datum.split(u' ')
            return contents[0]


# 从表格的4-7行中匹配出country
def find_country(sheet_param):
    for row_index in range(4, 7):
        row_content = sheet_param.row_values(row_index)
        for datum in row_content:
            if datum:
                contents = datum.split(u' ')
                if contents[1] == 'Metric' or contents[0] == 'WASDE':
                    break
                return contents[0]


# 找到market_year所在的行数下标
def find_market_year_line(sheet_param):
    p = re.compile(r'\d{4}\/\d{2}')
    # 在表格的8-10中寻找market_year行
    for i in range(8, 11):
        data = sheet_param.row_values(i)
        for datum in data:
            if datum:
                # 匹配时间-时间所在行即为market_year所在行
                if p.match(str(datum)[0:7]):
                    return i


# 从指定行中查询出attribute,如果不是attribute行，返回None
def find_attribute(sheet_param, row_index):
    for i in range(0, 3):
        content = sheet_param.cell(row_index, i).__str__()
        if content[0:5] != 'empty':
            return content[7:-1]
    return None


# 从指定行中查询出unit(保证指定的行一定包含unit)
def find_unit(sheet_param, row_index):
    row_content = sheet_param.row_values(row_index)
    for content_data in row_content:
        if content_data:
            return content_data


# 从market_year所在行起，寻找第一个空行，该空行即为page11中上半部分结束的位置
def find_first_end_index(begin_index, sheet_param):
    max_row_index = sheet_param.nrows - 1
    for i in range(begin_index, max_row_index):
        row_content = sheet_param.row_values(i)
        for content_data in row_content:
            if content_data:
                flag = False
                break
            else:
                flag = True
        if flag:
            return i


# 查询page 11下半部分的开始下标
def find_second_begin_index(sheet_param):
    max_row_index = sheet_param.nrows - 1
    count = 0
    for i in range(0, max_row_index):
        row_content = sheet_param.row_values(i)
        for datum in row_content:
            if datum:
                if str(datum).find('Supply and Use') != -1:
                    count += 1
                if count == 2:
                    return i


# 查询page 11下半部分的data部分开始下标
def find_second_data_begin_index(sheet_param, second_begin_index):
    max_row_index = sheet_param.nrows - 1
    count = 0
    for i in range(second_begin_index+1, max_row_index):
        # 第二个不为空的数据行的下标
        row_content = sheet_param.row_values(i)
        for datum in row_content:
            if datum:
                count += 1
                if count == 2:
                    return i
                else:
                    break


# 查询page 11下半部分的commodity_detail所在行的下标
def find_second_commodity_detail_index(sheet_param, second_begin_index):
    max_row_index = sheet_param.nrows - 1
    for i in range(second_begin_index+1, max_row_index):
        # 第一个不为空的数据行的下标
        row_content = sheet_param.row_values(i)
        for datum in row_content:
            if datum:
                return i


# 尝试获取下半部分的unit, 如果不是unit行，则返回None
def get_second_unit(sheet_param, row_index):
    row_content = sheet_param.row_values(row_index)
    p = re.compile(r'\d{4}\/\d{2}')
    unit_temp = ''
    count = 0
    flag = False
    for datum in row_content:
        if datum or datum == 0.0:
            count += 1
            if is_num(str(datum)[0]):
                flag = False
                break
            elif p.match(str(datum)[0:7]):
                flag = False
                break
            elif count >= 3:
                flag = False
                break
            else:
                flag = True
    # 是unit行，返回unit
    if flag:
        count_temp = 0
        for datum in row_content:
            if datum or datum == 0.0:
                count_temp += 1
                if count_temp == 1:
                    unit_temp = str(datum) + ' '
                else:
                    unit_temp += str(datum)
        return unit_temp
    else:
        return None


# 查询page 11中item数据，使用item_list返回
def find_data_page11(sheet_param, calendar_year_param):
    item_list = []
    month = find_default_month(sheet_param)
    commodity = find_commodity(sheet_param)
    country = find_country(sheet_param)
    # market_year所在行的下标
    market_year_line = find_market_year_line(sheet_param)
    # 表格上半部分结束下标
    first_end_index = find_first_end_index(market_year_line, sheet_param)
    # 表格数据正式开始的下标
    data_begin_index = market_year_line + 2

    # 上半部分的数据
    for index in range(data_begin_index, first_end_index):
        # 尝试查询出attribute
        attribute_temp = find_attribute(sheet_param, index)
        if attribute_temp:
            attribute = attribute_temp
        # 非attribute行,获取unit
        else:
            unit = find_unit(sheet_param, index)
            continue
        # 获取该行的数据
        row_content = sheet_param.row_values(index)
        j = 0
        for content_data in row_content:
            if content_data or content_data == 0.0:
                # 如果内容的第一个字符是数字，则表示是有效的value
                if is_num(str(content_data)[0:1]):
                    value = str(content_data)
                    # 获取market_year
                    market_year = sheet_param.cell(market_year_line, j).__str__()[7:-1]
                    # 尝试获取指定的month
                    month_content = sheet_param.cell(market_year_line+1, j).__str__()
                    if month_content[0:5] != 'empty' and month_content != 'text:u\'\'':
                        month = month_content[7:-1]
                    item = Item()
                    item.Commodity = commodity
                    item.Country = country
                    item.MarketYear = market_year
                    item.CalendarYear = calendar_year_param
                    item.Month = month
                    item.Attribute = attribute
                    item.Unit = unit
                    item.Value = value
                    item_list.append(item)
            j += 1

    # 下半部分的数据
    month = find_default_month(sheet_param)
    # 寻找page 11下半部分的开始下标
    second_begin_index = find_second_begin_index(sheet_param)
    country = find_second_country(sheet_param, second_begin_index)

    # commodity_detail 所在行的下标
    commodity_detail_index = find_second_commodity_detail_index(sheet_param, second_begin_index)
    # data开始的行数下标
    second_data_begin_index = find_second_data_begin_index(sheet_param, second_begin_index)
    max_row = sheet_param.nrows - 1
    market_year = ''
    attribute = ''
    p = re.compile(r'\d{4}\/\d{2}')
    for second_index in range(second_data_begin_index, max_row):
        # 尝试获取unit
        unit_temp = get_second_unit(sheet_param, second_index)
        if unit_temp:
            unit = unit_temp
        # 获取该行的数据
        row_content = sheet_param.row_values(second_index)
        valid_count = 2
        second_j = 0
        num_flag = False
        for content_data in row_content:
            if content_data or content_data == 0.0:
                # 尝试获取market_year
                if p.match(str(content_data)[0:7]):
                    market_year = str(content_data)
                # 尝试获取指定的month
                elif is_month(str(content_data)):
                    month = str(content_data)
                # 如果内容的第一个字符是数字，则表示是有效的value
                elif is_num(str(content_data)[0:1]):
                    num_flag = True
                    second_commodity = find_second_commodity(sheet_param, second_begin_index)
                    value = str(content_data)
                    # 尝试获取attribute
                    attribute_temp = sheet_param.cell(second_index, second_j - valid_count).__str__()
                    if attribute[0:5] != 'empty':
                        attribute = attribute_temp[7:-1]
                    # 尝试获取commodity_detail
                    commodity_detail = sheet_param.cell(commodity_detail_index, second_j).__str__()
                    if commodity_detail[0:5] != 'empty':
                        second_commodity = commodity_detail[7:-1] + '-' + second_commodity
                    item = Item()
                    item.Commodity = second_commodity
                    item.Country = country
                    item.MarketYear = market_year
                    item.CalendarYear = calendar_year_param
                    item.Month = month
                    item.Attribute = attribute
                    item.Unit = unit
                    item.Value = value
                    item_list.append(item)
            if num_flag:
                valid_count += 1
            second_j += 1
    return item_list


# 入口
if __name__ == "__main__":
    # # 单个文件测试
    # file_path = 'D:/data/test/test/wasde-07-11-2014.xls'
    # # 打开工作簿
    # work_book = xlrd.open_workbook(file_path)
    # sheet = work_book.sheet_param_by_name('Page 11')
    # calendar_year = get_calendar_year(file_path)
    # print find_second_country(sheet, 33)
    # print find_second_data_begin_index(sheet, 33)
    # print find_second_commodity_detail_index(sheet, 33)
    # print get_second_unit(sheet, 36)
    # print find_second_commodity(sheet, 33)
    # print find_default_month(sheet)
    # print find_commodity(sheet)
    # print find_market_year_line(sheet)
    # print find_attribute(sheet, 12)
    # print find_attribute(sheet, 13)
    # print find_unit(sheet, 13)
    # print find_country(sheet)
    # item_lists = find_first_data(sheet, calendar_year)
    # print find_second_begin_index(sheet)

    # 单个文件测试
    file_path = 'D:/data/wasde-12-10-2014.xls'
    # 打开工作簿
    work_book = xlrd.open_workbook(file_path)
    calendar_year = get_calendar_year(file_path)
    sheet = work_book.sheet_by_name('Page 11')
    item_lists = find_data_page11(sheet, calendar_year)
    for item in item_lists:
        print item.Country + ' ' + item.Commodity + '  ' + item.CalendarYear + '  ' + item.Attribute \
            + '  ' + item.Unit + '  ' + item.MarketYear + '  ' + item.Month + '  ' + item.Value