
from spreadsheets_lib import SheetReport 
import config as cfg
import get_time_working_off as report
from lib import *


report.getWorkTimeBalanceReport()
usersReport = report.users_work_time


def get_users_from_report():
   
    return [thing[0] for thing in usersReport]

def compareLists(user): 
    return user[0] not in users

def create_new_list_report():

    cur_day = get_cur_day()
    header = ["ФИО","Отгулял (ч)","Переработка (ч)","Отработано (ч)","Итог (ч)"]

    sheet_report.create_page(table_id, cur_day)
    sheet_report.add_line(table_id, cur_day, [header])
    allUsers = get_act_users()
    addUsers = list(filter(compareLists, allUsers))

    for adduser in addUsers:
        adduser.extend(["", "", "", 0])

    usersReport.extend(addUsers)

    usersReport.sort(key = lambda x: x[0])

    index = 2
    sheet_id = get_sheet_id(cur_day)
    for user_data in usersReport:
        sheet_report.add_line(table_id, cur_day , [user_data])

        if user_data[4] > 0:
            sheet_report.set_cell_format(table_id, sheet_id, index,  5, bg_color= {"red": 0, "green": 1, "blue": 0, "alpha": 1})
        elif user_data[4] == 0:
            sheet_report.set_cell_format(table_id, sheet_id, index,  5, bg_color={"red": 1, "green": 1, "blue": 1, "alpha": 1})
        else:
            sheet_report.set_cell_format(table_id, sheet_id, index, 5, bg_color={"red": 1, "green": 0, "blue": 0, "alpha": 1})
        index +=1 


users = get_users_from_report()       
create_new_list_report()






    