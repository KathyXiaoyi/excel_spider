# coding:utf-8
import pymongo
import xlrd
from util import get_file_list, save_item, if_need_deal, get_calendar_year, print_item
from spider_18_30 import match_data
from spider_11 import find_data_page11
from spider_12_15 import find_data_page12_15


# 入口
if __name__ == "__main__":
    con = pymongo.MongoClient('localhost', 27017)
    # 新建数据库
    my_db = con.USDA_WASDE
    # 新建table
    my_table = my_db.USDA_WASDE
    root_path = 'D:/data'
    file_list = get_file_list(root_path)
    for file_path in file_list:
        # 打开工作簿
        work_book = xlrd.open_workbook(file_path)
        calendar_year = get_calendar_year(file_path)
        print '-------------------------------------------------------------------------------------------------'
        print file_path
        sheet = work_book.sheet_by_name('Page 11')
        # page 11
        item_lists = find_data_page11(sheet, calendar_year)
        print 'Page 11'
        # print_item(item_lists)
        save_item(my_table, item_lists, con)

        for sheet_temp in work_book.sheets():
            if if_need_deal(sheet_temp, 12, 15):
                print '-------------------------------------'
                print sheet_temp.name
                # 将page 12-15 写入数据库
                item_lists = find_data_page12_15(sheet_temp, calendar_year)
                # print_item(item_lists)
                save_item(my_table, item_lists, con)

        for sheet_temp in work_book.sheets():
            if if_need_deal(sheet_temp, 18, 25):
                print '-------------------------------------'
                print sheet_temp.name
                # 将page 18-25写入数据库
                item_lists = match_data(sheet_temp, calendar_year)
                # print_item(item_lists)
                save_item(my_table, item_lists, con)
            if if_need_deal(sheet_temp, 28, 30):
                print '-------------------------------------'
                print sheet_temp.name
                # 将page 28-30写入数据库
                item_lists = match_data(sheet_temp, calendar_year)
                # print_item(item_lists)
                save_item(my_table, item_lists, con)
