
from spreadsheets_lib import SheetReport 
import config as cfg
import get_time_working_off as report
from lib import *
import time

report.getWorkTimeBalanceReport()
usersReport = report.users_work_time


def get_users_from_report():
   
    return [thing[0] for thing in usersReport]

def compareLists(user): 
    return user[0] not in users

def create_new_list_report():

    get_cur_m = today_date()
    header = ["ФИО","Отгулял (ч)","Переработка (ч)","Отработано (ч)","Итог (ч)"]

    sheet_report.create_page(table_id, get_cur_m)
    sheet_report.add_line(table_id, get_cur_m, [header])
    allUsers = get_act_users()
    addUsers = list(filter(compareLists, allUsers))

    for adduser in addUsers:
        adduser.extend(["", "", "", 0])

    usersReport.extend(addUsers)

    usersReport.sort(key = lambda x: x[0])

    index = 2
    sheet_id = get_sheet_id(get_cur_m)


    sheet_report.add_line(table_id, get_cur_m , usersReport)

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
users = get_users_from_report()       
create_new_list_report()






    