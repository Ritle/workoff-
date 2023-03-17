from spreadsheets_lib import SheetReport 

import get_time_working_off as report
from lib import *



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

def compareLists(user, users): 
    return user[0] not in users

def check_itog(itog):
    return float(itog[4]) != 0.0

def fill_final_report(list_name, month_name, col):

    
    print(f"Заполняю сводную за {col[0]}")
    num_len = len(sheet_report.read_list(table_id, "Сводная")["values"][0])

    if num_len > 1:
        last_month = col

        pre_last_month = get_last_stat_sheet()
        new_list = []
        for (last, pre_last) in zip(last_month ,pre_last_month ):
            if last == list_name:
                new_list.append(last)
            else:   
                new_list.append(float(last) + float(pre_last.replace(',', '.')))
        col = new_list

    data_set = []
    
    convert_arr = convertDate(col)

    for convert_arr_v in convert_arr:
        data_set.append([convert_arr_v])

    col_num = getIndexMonth(month_name) + 1
    sheet_report.insert_range(table_id, "Сводная", data_set, 0, len(data_set), col_num, col_num)
    sheet_report.autoSizeColumn(table_id, get_sheet_id("Сводная"), 0, 100)

def create_new_list_report(users,month_name, usersReport, list_name):

    
    header = ["Сотрудник учебного центра","Отгулял (ч)","Переработка (ч)","Отработано (ч)","Итог (ч)"]
    
    sheet_report.create_page(table_id, list_name)

    id_list = get_sheet_id(list_name)
    
    sheet_report.set_format(table_id, id_list, 4, 5, "NUMBER_GREATER", {"red": 0.5, "green": 0.8, "blue": 0.5, "alpha": 1})
    sheet_report.set_format(table_id, id_list, 4, 5, "NUMBER_LESS", {"red": 0.9, "green": 0.5, "blue": 0.5, "alpha": 1})

    sheet_report.add_line(table_id, list_name, [header])
    sheet_report.autoSizeColumn(table_id, id_list, 0, 5)
    allUsers = get_act_users()
   
    addUsers = [element for element in allUsers if compareLists(element, users)]
    for adduser in addUsers:
        adduser.extend(["", "", "", "0.0"])

    usersReport.extend(addUsers)
    usersReport.sort(key = lambda x: x[0])

    new_c = []
    for c_v in usersReport:
        new_elem = convertDate(c_v)
        new_c.append(new_elem)

    sheet_report.add_line(table_id, list_name , new_c)
    final_report = [list_name]

    final_report.extend([thing[4] for thing in usersReport])

    fill_final_report(list_name, month_name, final_report)

    print("Готово!")


def start(sid):

    for month in months:

        sheets = sheet_report.sheet_list(table_id)
        month_num = int(month.split("-")[1][1])
        year_num = month.split("-")[0]
    
        month_name = convertNumMonth(month_num)
        list_name = get_last_day_month(year_num, month_num)
        print(f"Собираю статистику по  {list_name}")      
        if list_name not in sheets:

            user_report = report.getWorkTimeBalanceReport(sid, month)
            users = get_users_from_report(user_report)  
            
            print(f"Создаю новый лист {list_name}")    
            create_new_list_report(users, month_name, user_report, list_name)
        else:
            print(f"Статистика за {month_name} {year_num} уже собрана") 

       



    