from spreadsheets_lib import SheetReport 
import config as cfg
import get_time_working_off as report
from lib import *
import time

def convertNumMonth(num):
    
    return calendar[num - 1]

def getIndexMonth(name):
    return calendar.index(name)

def get_users_from_report(usersReport):
   
    return [thing[0] for thing in usersReport]

def compareLists(user): 
    return user[0] not in users

def fill_final_report(col):
    print(f"Заполняю сводную за {col[0]}")

    data_set = []
    
    for col_v in col:
        data_set.append([col_v])

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

    index = 2
    sheet_id = get_sheet_id(list_name)


    sheet_report.add_line(table_id, list_name , usersReport)

    for user_data in usersReport:
        print(f"Проверяем итог у {user_data[0]}")
        if user_data[4] > 0:
            sheet_report.set_cell_format(table_id, sheet_id, index,  5, bg_color= {"red": 0.5, "green": 0.8, "blue": 0.5, "alpha": 1})

        if user_data[4] < 0:
            sheet_report.set_cell_format(table_id, sheet_id, index,  5, bg_color= {"red": 0.9, "green": 0.5, "blue": 0.5, "alpha": 1})
        index +=1 
        time.sleep(1)

    final_report = [list_name]

    final_report.extend([thing[4] for thing in usersReport])

    
    print("Готово!")

    return final_report
    
months = get_need_months()

today_date_v = today_date() 
final_report = []
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
        final_report.append(create_new_list_report(user_report, list_name))  
    else:
        print(f"Статистика за {month_name} {year_num} уже собрана") 


    #if len(final_report) > 1:
        






    