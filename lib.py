from datetime import datetime, date, timedelta
from calendar import monthrange
from spreadsheets_lib import SheetReport 
import config as cfg
import traceback


exceptionUsers = ["Гордеев Андрей", "Липовская Наталья", "Лутченко Юлия", "Макина Кристина", "Соловьева Екатерина", "Шарова Екатерина"]
calendar = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
table_id = cfg.google_report_table_id


sheet_report = SheetReport(cfg.token_str, is_dev=False) # создаем объект который читает из экселя

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

def get_last_day_month(year, month):
    last_day = monthrange(int(year), month)[1]
    return f"{last_day}/{month}/{year}"

def get_last_stat_sheet():

    fullstat = sheet_report.read_list(table_id, "Сводная")["values"]
    
    return [str_to_hours(full[-1]) for full in fullstat]


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

def last_day_of_month(any_day):
    next_month = any_day.replace(day=28) + datetime.timedelta(days=4)  # this will never fail
    return next_month - datetime.timedelta(days=next_month.day)

def convertDate(c):
    new = [] 
    for c_v in c:
        try:
            c_v = str(c_v)
            arr_c = c_v.split('.')
            if(int(arr_c[1]) == 0):

                if(c_v[0] == "-"):
                    t = f"-{int(abs(int(arr_c[0])))}:00:00"
                else:
                    t = f"{int(abs(int(arr_c[0])))}:00:00"
                
                new.append(t)
            elif (int(arr_c[1]) == 1):
                hour = int(arr_c[0])

                if(c_v[0] == "-"):
                    t = f"-{abs(hour)}:0{abs(round((float(c_v)-hour)*60))}:00"   
                else:
                    t = f"{abs(hour)}:0{abs(round((float(c_v)-hour)*60))}:00" 

                new.append(t)
            elif (int(arr_c[1]) > 1):
                hour = int(arr_c[0])

                if(c_v[0] == "-"):
                    t = f"-{abs(hour)}:{abs(round((float(c_v)-hour)*60))}:00"   
                else:
                    t = f"{abs(hour)}:{abs(round((float(c_v)-hour)*60))}:00" 

                new.append(t)
            elif(c_v == ""): 
                new.append(c_v)
            elif(c_v == '0.0'): 
                new.append("0:00:00")
        except Exception:
            print('Ошибка:\n', traceback.format_exc())
            new.append(c_v)
    return new


def str_to_hours(s):
    arr_s = s.split(':')

    if '/' not in s:

        if(s[0] == '-'):

            return f"-{abs(int(arr_s[0]))}.{int(int(arr_s[1])/6)}"
        else:
            return f"{abs(int(arr_s[0]))}.{int(int(arr_s[1])/6)}"
    
months = get_need_months()
today_date_v = today_date() 

