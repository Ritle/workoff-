from datetime import datetime, date, timedelta
from spreadsheets_lib import SheetReport 
import config as cfg


exceptionUsers = ["Гордеев Андрей", "Липовская Наталья", "Лутченко Юлия", "Макина Кристина", "Соловьева Екатерина", "Шарова Екатерина"]
calendar = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
table_id = cfg.google_report_table_id


sheet_report = SheetReport(cfg.token_str, is_dev=True) # создаем объект который читает из экселя

def today_date():
    return str(datetime.now()).split()[0]

def get_cur_month():
    return str(datetime.today().replace(day=1)).split()[0]

def get_last_month(): 

    last_day_of_prev_month = date.today().replace(day=1) - timedelta(days=1)
    return str(date.today().replace(day=1) - timedelta(days=last_day_of_prev_month.day))

def get_cur_day():
    return str(datetime.today()).split(".")[0].replace(':', '-')

def get_act_users():
    
    users = sheet_report.read_list(table_id, 'Список сотрудников')['values']
    del users[0]
    return users


def get_sheet_id(sheet_name):
    sheets = sheet_report.sheet_list(table_id)

    return sheets[sheet_name]

def get_need_months():

    today = int(get_cur_day().split()[0].split('-')[1][1])

    year =  int(get_cur_day().split()[0].split('-')[0])

    arr_first = []

    for month in range(1, today + 1):

        if month < 10:
            month = "0" + str(month) 
        date_str = f"{year}-{month}-01"
        arr_first.append(date_str)

    return arr_first



def get_last_stat_sheet():

    fullstat = sheet_report.read_list(table_id, "Сводная")["values"]

    

    return [full[-1] for full in fullstat]


def findIndexNameInReport(name, report):
    for v in report:
        if name in v:
            return report.index(v)


def convertNumMonth(num):
    
    return calendar[num - 1]

def getIndexMonth(name):
    return calendar.index(name)

def get_users_from_report(usersReport):
   
    return [thing[0] for thing in usersReport]

def convertDate(c):

    try:
        c = str(c)
        arr_c = c.split('.')
        if(int(arr_c[1]) == 0):
            t = f"{int(c[0])}:00"
        if (int(arr_c[1]) > 0):
            hour = int(arr_c[0])
            t = f"{hour}:{abs(round((float(c)-hour)*60))}"    
            return t
        elif(c == ""): 
            return c
        elif(c == '0.0'): 
            return f"0:00"
    except:
        return c

