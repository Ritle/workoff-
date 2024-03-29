import json
from transport import RPCCall
from datetime import datetime as dtm
from datetime import timedelta as tdm
import time
import config as cfg

import lib 
 
 #объект который делает вызов в сбис


def getWorkTimeBalanceReport(user_sid , month):

    sbis_client = RPCCall(site="main", base="online.sbis.ru", sid=user_sid)
    users_work_time = []
    page_num = 0
    users_report_json = {"jsonrpc":"2.0","protocol":6,"method":"Отгул.WorkTimeBalanceReport",
    "params":{"Фильтр":{"d":[None, month,"-2",14466637,None,None,0],
    "s":[{"t":"Строка","n":"Parent"},{"t":"Строка","n":"Дата"},{"t":"Строка","n":"Организация"},{"t":"Число целое","n":"Подразделение"},{"t":"Строка","n":"Состояния"},{"t":"Строка","n":"ТипОтгула"},{"t":"Число целое","n":"ТипОтчета"}],"_type":"record","f":0},"Сортировка":None,
    "Навигация":{"d":[True,25,page_num],"s":[{"t":"Логическое","n":"ЕстьЕще"},{"t":"Число целое","n":"РазмерСтраницы"},{"t":"Число целое","n":"Страница"}],"_type":"record","f":0},"ДопПоля":[]},"id":1}
    users = sbis_client.post_request(users_report_json).recordset()
    while users:
        for user in users:
            if user["Name"] not in lib.exceptionUsers:
                user_name_value = user["Name"]
                user_PK_value = user["PK"]
                user_PersonID_value = user["PersonID"]
                workedOffTime = user["TimeoffTime"]
                users_work_time_users = getUserReport(sbis_client, user_name_value, user_PK_value, user_PersonID_value, workedOffTime, month)
                users_work_time.append(users_work_time_users)
        page_num += 1
        users_report_json = {"jsonrpc":"2.0","protocol":6,"method":"Отгул.WorkTimeBalanceReport",
        "params":{"Фильтр":{"d":[None, month,"-2",14466637,None,None,0],
        "s":[{"t":"Строка","n":"Parent"},{"t":"Строка","n":"Дата"},{"t":"Строка","n":"Организация"},{"t":"Число целое","n":"Подразделение"},{"t":"Строка","n":"Состояния"},{"t":"Строка","n":"ТипОтгула"},{"t":"Число целое","n":"ТипОтчета"}],"_type":"record","f":0},"Сортировка":None,
        "Навигация":{"d":[True, 25, page_num],"s":[{"t":"Логическое","n":"ЕстьЕще"},{"t":"Число целое","n":"РазмерСтраницы"},{"t":"Число целое","n":"Страница"}],"_type":"record","f":0},"ДопПоля":[]},"id":1}
        users = sbis_client.post_request(users_report_json).recordset()

    return users_work_time

def getUserReport(sbis_client, user_name, user_PK, user_PersonID, workedOffTime, month):

    time.sleep(1)
    print("Проверяем", user_name)

    page_num_user = 0
    workTime = 0
    workTimeOut = 0
    WorkTimeBalanceReportUser = {"jsonrpc":"2.0","protocol":6,"method":"Отгул.WorkTimeBalanceReport",
    "params":{"Фильтр":{"d":[user_PK,month,"-2",user_PersonID,None,None,0],"s":[{"t":"Число целое","n":"Parent"},{"t":"Строка","n":"Дата"},{"t":"Строка","n":"Организация"},{"t":"Число целое","n":"Подразделение"},{"t":"Строка","n":"Состояния"},{"t":"Строка","n":"ТипОтгула"},{"t":"Число целое","n":"ТипОтчета"}],"_type":"record","f":0},"Сортировка":None,
    "Навигация":{"d":[True,25,page_num_user],"s":[{"t":"Логическое","n":"ЕстьЕще"},{"t":"Число целое","n":"РазмерСтраницы"},{"t":"Число целое","n":"Страница"}],"_type":"record","f":0},"ДопПоля":[]},"id":1}
    user_report = sbis_client.post_request(WorkTimeBalanceReportUser).recordset()

    while user_report:
    
        for report_work_time in user_report:
            if report_work_time["Name"] != "Отработано": 
                if report_work_time["Details"]:
                    if "TotalTime" in report_work_time["Details"][0]:
                        workTime += report_work_time["Details"][0]["TotalTime"]        
            else:
                workTimeOut += report_work_time["WorkedOffTime"]

        page_num_user += 1
        WorkTimeBalanceReportUser = {"jsonrpc":"2.0","protocol":6,"method":"Отгул.WorkTimeBalanceReport",
            "params":{"Фильтр":{"d":[user_PK,month,"-2",user_PersonID,None,None,0],"s":[{"t":"Число целое","n":"Parent"},{"t":"Строка","n":"Дата"},{"t":"Строка","n":"Организация"},{"t":"Число целое","n":"Подразделение"},{"t":"Строка","n":"Состояния"},{"t":"Строка","n":"ТипОтгула"},{"t":"Число целое","n":"ТипОтчета"}],"_type":"record","f":0},"Сортировка":None,
            "Навигация":{"d":[True,25,page_num_user],"s":[{"t":"Логическое","n":"ЕстьЕще"},{"t":"Число целое","n":"РазмерСтраницы"},{"t":"Число целое","n":"Страница"}],"_type":"record","f":0},"ДопПоля":[]},"id":1}

        new_user_rep = sbis_client.post_request(WorkTimeBalanceReportUser)
        new_user_rep_json = new_user_rep.response.json()

        if new_user_rep_json["result"]["n"]:

            user_report = new_user_rep.recordset()
        else:
            user_report = None

    if workTime == None:
        workTime = 0
    if workedOffTime == None:
        workedOffTime = 0
    if workTimeOut == None:
        workTimeOut = 0

    itog = round (((workTime + workTimeOut) - workedOffTime)/60, 1)

    if workTime != 0:
        workTime = round(workTime/60, 1)
    else:
        workTime = ""

    if workedOffTime and workedOffTime != 0:
        workedOffTime = round(workedOffTime/60, 1)
    else:
        workedOffTime = ""

    if workTimeOut and workTimeOut != 0:
        workTimeOut = round(workTimeOut/60, 1 )
    else:
        workTimeOut = ""
   
    print( f"У {user_name}, отгулов {workedOffTime} ч. переработок {workTime} ч. отработано календарем {workTimeOut} ч") 

    return [user_name, workedOffTime, workTime,  workTimeOut, itog]



