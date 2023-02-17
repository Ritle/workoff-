from datetime import datetime
from spreadsheets_lib import SheetReport 
import config as cfg


table_id = cfg.google_report_table_id


sheet_report = SheetReport(cfg.token_str, is_dev=False) # создаем объект который читает из экселя

def get_cur_month():
    return str(datetime.today().replace(day=1)).split()[0]

def get_cur_day():

    
    return str(datetime.today()).split(".")[0].replace(':', '-')

def get_act_users():
    
    users = sheet_report.read_list(table_id, 'Список сотрудников')['values']
    del users[0]
    return users


def get_sheet_id(sheet_name):
    sheets = sheet_report.sheet_list(table_id)

    return sheets[sheet_name]
