# coding:utf-8
from util import month_map, is_num, get_calendar_year, is_month, get_file_list, is_date
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
def find_commodity(sheet_param, market_year_line):
    row_content = sheet_param.row_values(market_year_line)
    for datum in row_content:
        if datum:
            return datum


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
    # 在表格的6-11中寻找market_year行
    for i in range(6, 11):
        data = sheet_param.row_values(i)
        for datum in data:
            if datum or data == 0.0:
                # 匹配时间-时间所在行即为market_year所在行
                if p.match(str(datum)[0:7]):
                    return i


# 匹配出commodity所在的行
def if_new_commodity_line(sheet_param, row_index):
    p = re.compile(r'\d{4}\/\d{2}')
    data = sheet_param.row_values(row_index)
    for datum in data:
        if datum or data == 0.0:
            # 匹配时间-时间所在行即为commodity所在行
            if p.match(str(datum)[0:7]):
                return True
    return False


# 从指定行中查询出attribute,如果不是attribute行，返回None
def find_attribute(sheet_param, row_index):
    for i in range(0, 5):
        content = sheet_param.cell(row_index, i).__str__()
        if content[0:5] != 'empty':
            # 判断它的上两行，如何开始是时间，表示是unit行，返回None
            content_top = sheet_param.cell(row_index-2, i).__str__()
            content_top2 = sheet_param.cell(row_index - 3, i).__str__()
            if content_top or content_top2:
                if is_date(content_top[7:14]) or is_date(content_top2[7:14]):
                    return None
                else:
                    return content[7:-1]
    return None


# 从指定行中查询出unit(保证指定的行一定包含unit)
def find_unit(sheet_param, row_index):
    row_content = sheet_param.row_values(row_index)
    count = 1
    content_temp = ''
    for content_data in row_content:
        if content_data and content_data != 'Filler' and not is_num(content_data) and not is_month(content_data[0:3]):
            if count == 1:
                content_temp = content_data
                count = 2
            else:
                content_temp = content_temp + ' ' + content_data
    return content_temp


# 查询page 12-15中item数据，使用item_list返回
def find_data_page12_15(sheet_param, calendar_year_param):
    p = re.compile(r'\d{4}\/\d{2}')
    max_row_index = sheet_param.nrows - 1
    item_list = []
    month = find_default_month(sheet_param)
    country = find_country(sheet_param)
    # market_year所在行的下标
    market_year_line = find_market_year_line(sheet_param)
    # 查找出第一个commodity
    commodity = find_commodity(sheet_param, market_year_line)
    # 表格结束下标
    first_end_index = max_row_index
    # 表格数据正式开始的下标
    data_begin_index = market_year_line + 2
    attribute = ''
    unit = ''
    # 数据部分
    for index in range(data_begin_index, first_end_index):
        # 尝试查询出attribute
        attribute_temp = find_attribute(sheet_param, index)
        if attribute_temp and attribute_temp != 'Filler':
            attribute = attribute_temp
        # 非attribute行,获取unit
        else:
            unit_temp = find_unit(sheet_param, index)
            if unit_temp is None or unit_temp.strip() == '':
                pass
            else:
                unit = unit_temp
        # 获取该行的数据
        row_content = sheet_param.row_values(index)
        if if_new_commodity_line(sheet_param, index):
            for content_data in row_content:
                if content_data:
                    commodity = content_data
                    break
        j = 0
        for content_data in row_content:
            if content_data or content_data == 0.0:
                # 如果内容的第一个字符是数字且不是时间，则表示是有效的value
                if is_num(str(content_data)[0:1]) and not p.match(str(content_data)[0:7]):
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
                    item.Value = str(value)
                    item_list.append(item)
            j += 1
    return item_list


# 入口
if __name__ == "__main__":
    # 单个文件测试
    file_path = 'D:/data/wasde-12-09-2015.xls'
    # 打开工作簿
    work_book = xlrd.open_workbook(file_path)
    calendar_year = get_calendar_year(file_path)
    sheet = work_book.sheet_by_name('Page 12')
    item_lists = find_data_page12_15(sheet, calendar_year)
    for item in item_lists:
        print item.Country + ' ' + item.Commodity + '  ' + item.CalendarYear + '  ' + item.Attribute \
            + '  ' + item.Unit + '  ' + item.MarketYear + '  ' + item.Month + '  ' + item.Value

    # calendar_year = get_calendar_year(file_path)
    # print find_default_month(sheet)
    # print find_commodity(sheet)
    # print find_market_year_line(sheet)
    # print find_attribute(sheet, 12)
    # print find_attribute(sheet, 13)
    # print find_unit(sheet, 11)
    # print find_country(sheet)

    # # 多个文件测试
    # root_path = 'D:/data'
    # file_list = get_file_list(root_path)
    # for file_path in file_list:
    #     calendar_year = get_calendar_year(file_path)
    #     # 打开工作簿
    #     work_book = xlrd.open_workbook(file_path)
    #     sheet = work_book.sheet_by_name('Page 15')
    #     calendar_year = get_calendar_year(file_path)
    #     item_lists = find_data_page12_15(sheet, calendar_year)
    #     for item in item_lists:
    #         print item.Country + ' ' + item.Commodity + '  ' + item.CalendarYear + '  ' + item.Attribute \
    #             + '  ' + str(item.Unit) + '  ' + item.MarketYear + '  ' + item.Month + '  ' + item.Value
        #print find_default_month(sheet_param)
        #print find_commodity(sheet_param)
        #print find_market_year_line(sheet_param)
        #begin = find_market_year_line(sheet_param)
        # print find_first_end_index(begin, sheet_param)
