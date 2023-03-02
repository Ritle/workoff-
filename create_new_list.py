from spreadsheets_lib import SheetReport 
import config as cfg
import get_time_working_off as report
from lib import *
import time


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

def compareLists(user): 
    return user[0] not in users

def check_itog(itog):
    return itog[4] != 0

def fill_final_report(col):

    print(f"Заполняю сводную за {col[0]}")

    num_len = len(sheet_report.read_list(table_id, "Сводная")["values"][0])

    if num_len > 1:
        last_month = col
        pre_last_month = get_last_stat_sheet()
        new_list = []
        for (last, pre_last) in zip(last_month ,pre_last_month ):
            if type(last) == str and type(pre_last) ==str:
                new_list.append(last)
            else:   
                new_list.append(last + float(pre_last.replace(',', '.')))
        col = new_list

    data_set = []
    
    for col_v in col:
        data_set.append([convertDate(col_v)])

    col_num = getIndexMonth(col[0].split()[0])
    sheet_report.insert_range(table_id, "Сводная", data_set, 0, len(data_set), col_num+1, col_num+1)

def create_new_list_report(usersReport, list_name):
    header = ["ФИО","Отгулял (ч)","Переработка (ч)","Отработано (ч)","Итог (ч)"]

    sheet_report.create_page(table_id, list_name)
    sheet_report.add_line(table_id, list_name, [header])
    allUsers = get_act_users()
    addUsers = list(filter(compareLists, allUsers))

    for adduser in addUsers:
        adduser.extend(["", "", "", 0])

    usersReport.extend(addUsers)
    usersReport.sort(key = lambda x: x[0])

    new_c = []
    for c_v in usersReport:
        new_c.append(convertDate(c_v))

    sheet_id = get_sheet_id(list_name)
    sheet_report.add_line(table_id, list_name , new_c)

    format_user_report = list(filter(check_itog, usersReport))

    for user_data in format_user_report:
        print(f"Проверяем итог у {user_data[0]}")
        index = findIndexNameInReport(user_data[0], usersReport) +  2
        if user_data[4] > 0:
            sheet_report.set_cell_format(table_id, sheet_id, index,  5, bg_color= {"red": 0.5, "green": 0.8, "blue": 0.5, "alpha": 1})

        if user_data[4] < 0:
            sheet_report.set_cell_format(table_id, sheet_id, index,  5, bg_color= {"red": 0.9, "green": 0.5, "blue": 0.5, "alpha": 1})
       
        time.sleep(1)

    final_report = [list_name]

    final_report.extend([thing[4] for thing in usersReport])

    fill_final_report(final_report)

    print("Готово!")

    
months = get_need_months()

today_date_v = today_date() 

for month in months:

    sheets = sheet_report.sheet_list(table_id)
    month_num = int(month.split("-")[1][1])
    year_num = month.split("-")[0]

    month_name = convertNumMonth(month_num)
    list_name = f"{month_name} {year_num} {today_date_v}"
    print(f"Собираю статистику по  {list_name}")      
    if list_name not in sheets:

        user_report = report.getWorkTimeBalanceReport(month)
        users = get_users_from_report(user_report)  
        
        print(f"Создаю новый лист {list_name}")    
        create_new_list_report(user_report, list_name)
    else:
        print(f"Статистика за {month_name} {year_num} уже собрана") 

       






    