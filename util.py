# coding:utf-8
import os
import re


# 判断字段是不是月份
def is_month(content_param):
    month_list = ['Jan', 'Feb', 'Mar',  'Apr', 'May', 'Jun',  'Jul', 'Aug', 'Sep', 'Oct', 'Nov',  'Dec']
    for month in month_list:
        if content_param == month:
            return True
    return False


# 月份简写对应关系
def month_map(key):
    dict_map = {'January': 'Jan', 'February': 'Feb', 'March': 'Mar', 'April': 'Apr',
                'May': 'May', 'June': 'Jun', 'July': 'Jul', 'August': 'Aug', 'September': 'Sep', 'October': 'Oct',
                'November': 'Nov', 'December': 'Dec'}
    return dict_map.get(key)


# 根据文件名 获取calendar_year
def get_calendar_year(file_path):
    suffix_index = file_path.find('.')
    suffix = file_path[suffix_index+1:]
    if suffix == 'xls':
        return file_path[-8:-4]
    else:
        return file_path[-9:-5]


# 获取根路径下的所有文件路径
def get_file_list(root_path_param):
    file_lists = []
    file_names = os.listdir(root_path_param)
    if len(file_names) > 0:
        for fn in file_names:
            # 默认直接返回所有文件名
            full_file_name = root_path_param+'/'+fn
            file_lists.append(full_file_name)
    return file_lists


# 自定义方法获取自增ID
def get_id(con):
    # 新建数据库
    my_db = con.ids
    # 新建table
    my_table = my_db.id
    my_table.find_and_modify({"_id": "test"}, update={"$inc": {'count': 1}}, upsert=True)
    return my_table.find_one({"_id": "test"})['count']


# 将item列表保存到数据库中
def save(my_table, lists, con):
    for item_list in lists:
        save_item(my_table, item_list, con)


def save_item(my_table, lists, con):
    for item in lists:
        if item.Value == 'empty' or item.Value == '' or item.Value == 'NA' or item.Value == 'NA\'':
            continue
        else:
            # 自增ID
            item._id = get_id(con)
            item_dict = item.__dict__
            my_table.save(item_dict)


# 判断该sheet是否是需要处理的sheet
def if_need_deal(sheet_param, begin_index, end_index):
    page_name = sheet_param.name
    page_name = page_name.split(' ')
    if page_name[1].isdigit():
        page_name[1] = int(page_name[1])
        if begin_index <= page_name[1] <= end_index:
            return True
    else:
        return False


# 判断数据是否是数字
def is_num(datum):
    flag = False
    try:
        float(datum)
        flag = True
    except ValueError:
        flag = False
    finally:
        return flag


# 判断是否是日期
def is_date(content):
    p = re.compile(r'\d{4}\/\d{2}')
    return p.match(str(content))

# 打印item_list
def print_item(item_list):
    for item in item_list:
        print str(item.Country) + ' ' + str(item.Commodity) + '  ' + str(item.CalendarYear) + '  ' + str(item.Attribute) \
              + '  ' + str(item.Unit) + '  ' + str(item.MarketYear) + '  ' + str(item.Month) + '  ' + str(item.Value)