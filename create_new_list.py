from spreadsheets_lib import SheetReport 
import config as cfg
import get_time_working_off as report
from lib import *
import time


def get_users_from_report(usersReport):
   
    return [thing[0] for thing in usersReport]

def compareLists(user): 
    return user[0] not in users

def create_new_list_report(usersReport, list_name):

    
    header = ["ФИО","Отгулял (ч)","Переработка (ч)","Отработано (ч)","Итог (ч)"]

    sheet_report.create_page(table_id, month + " "+ list_name)
    sheet_report.add_line(table_id, list_name, [header])
    allUsers = get_act_users()
    addUsers = list(filter(compareLists, allUsers))

    for adduser in addUsers:
        adduser.extend(["", "", "", 0])

    usersReport.extend(addUsers)

    usersReport.sort(key = lambda x: x[0])

    index = 2
    sheet_id = get_sheet_id(list_name)
    sheet_report.add_line(table_id, list_name , usersReport)

    for user_data in usersReport:
    
        if user_data not in exceptionUsers:

            print(f"Проверяем итог у {user_data[0]}")

            if user_data[4] > 0:
                sheet_report.set_cell_format(table_id, sheet_id, index,  5, bg_color= {"red": 0.2, "green": 0.6, "blue": 0.2, "alpha": 0.8})
            elif user_data[4] == 0:
                sheet_report.set_cell_format(table_id, sheet_id, index,  5, bg_color={"red": 1, "green": 1, "blue": 1, "alpha": 1})
            else:
                sheet_report.set_cell_format(table_id, sheet_id, index, 5, bg_color={"red": 0.75, "green": 0.15, "blue": 0.1, "alpha": 0.8})
            index +=1 
        time.sleep(1)
    print("Готово!")


months = get_need_months()

today_date_v = today_date() 

for month in months:

    print(f"Собираю статистику по  {month} в {today_date_v}")
    
    user_report = report.getWorkTimeBalanceReport(month)
    users = get_users_from_report(user_report)  
    list_name = f"{month} {today_date_v}" 
    print(f"Создаю новый лист {list_name}")    
    create_new_list_report(user_report, list_name)





    