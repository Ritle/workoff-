__author__ = 'av.bardashov'

# ---------------------------------------------------------------------------------------------------------------------
# Модуль для работы с документами google spreadsheets, в которых ведется аназиз упавших сборок
# - создает новый док каждый месяц
# - Добавляет запись о упавшей сборке в гугл докс
# - Добавляет комментарий к сборке
# - Создает док с текстом соообщения для отправки в телегу
# ---------------------------------------------------------------------------------------------------------------------
import base64
import pickle
import socket
import time

from googleapiclient import discovery


class SheetReport():
    """
    Инструменты для работы google таблицами
    """
    # token_byte = open('token.pickle', 'rb').read()
    # token_str = base64.b64encode(token_byte).decode('utf-8')
    link_temple = "https://docs.google.com/spreadsheets/d/{}"

    def __init__(self, token, is_dev=True):
        """
        Конструктор класса
        При получении параметра token создаётся объект sheet_obj, который позволяет работать с таблицей.
        :param token: Токен для авторизации в google передается строкой
        """ 
        socket.setdefaulttimeout(600)
        if is_dev:
            self.sheet_obj = discovery.build('sheets', 'v4', developerKey=token).spreadsheets()
        else:
            creds = pickle.loads(base64.b64decode(token))
            self.sheet_obj = discovery.build('sheets', 'v4', credentials=creds).spreadsheets()
        
    def sheet_list(self, sheet_id):
        """
        Получаем список листов, их Id и название. Одновременно список помещается в переменную self.list_sheets
        :return: Список листов в формате словаря {"id": ..., "name": ...}
        """
        err_cnt = 5
        err_msg = ''
        while err_cnt > 0:
            try:
                spreadsheet = self.sheet_obj.get(spreadsheetId=sheet_id).execute()
                sheetList = spreadsheet.get('sheets')
                # return [{"id": it['properties']['sheetId'], "name": it['properties']['title']} for it in sheetList]
                return {it['properties']['title']: it['properties']['sheetId'] for it in sheetList}
            except Exception as e:
                err_msg = str(e)
                print(f'{err_cnt}: {err_msg}')
                err_cnt -= 1
                time.sleep(60)
        print(f'Ошибка при отправке запроса sheet_list():\n{err_msg}', 'e')
        exit(1)
        
    def read_list(self, sheet_id, list_name):
        """ """
        err_cnt = 5
        err_msg = ''
        while err_cnt > 0:
            try:
                return self.sheet_obj.values().get(spreadsheetId=sheet_id, range=list_name, 
                        valueRenderOption='FORMATTED_VALUE', dateTimeRenderOption='SERIAL_NUMBER').execute()
            except Exception as e:
                err_msg = str(e)
                print(f'{err_cnt}: {err_msg}')
                err_cnt -= 1
                time.sleep(60)
        print(f'Ошибка при отправке запроса read_list(..,{list_name}):\n{err_msg}', 'e')
        exit(1)

    def create_doc(self, doc_name):
        """ """
        req_body = {"properties": {"title": doc_name}}
        new_sheet = self.sheet_obj.create(body=req_body, fields='spreadsheetId').execute()
        new_id = new_sheet.get('spreadsheetId')
        print(f'Создали новый документ {self.link_temple.format(new_id)}')
        return new_id

    def copy_data(self, last_sheet_id, new_sheet_id):
        """ """
        tmpl_sheets = self.sheet_obj.get(spreadsheetId=last_sheet_id).execute().get('sheets')
        rename_request_body = {"requests": [{"updateSheetProperties": {"properties": {"sheetId": 0,"title": ""},"fields": "title"}}]}
        copy_sheet_request_body = {'destination_spreadsheet_id':new_sheet_id}

        for tmpl_sheet in tmpl_sheets:
            # создаем новый лист и копируем в него данные из шаблона
            new_sheet = self.sheet_obj.sheets().copyTo(spreadsheetId=last_sheet_id, sheetId=tmpl_sheet['properties']['sheetId'], 
                                                        body=copy_sheet_request_body).execute()
            # переименовываем лист
            rename_request_body['requests'][0]['updateSheetProperties']['properties']['sheetId'] = new_sheet['sheetId']
            rename_request_body['requests'][0]['updateSheetProperties']['properties']['title'] = tmpl_sheet['properties']['title']
            self.sheet_obj.batchUpdate(spreadsheetId=new_sheet_id, body=rename_request_body).execute()
            # Очищаем все поля
            if not tmpl_sheet['properties']['title'] in ['Пояснения', 'Проблемы']:
                self.clear_list(new_sheet_id, tmpl_sheet['properties']['title'], new_sheet['gridProperties']['rowCount'])
            print(f"Скопировали лист {tmpl_sheet['properties']['title']}")

    def del_list(self, sheet_id, list_id):
        """ """
        self.sheet_obj.batchUpdate(spreadsheetId=sheet_id,body={"requests":[{"deleteSheet":{"sheetId": list_id}}]}).execute()

    def clear_list(self, sheet_id, list_title, rows_cnt):
        """ 
        rows_cnt = new_sheet['gridProperties']['rowCount']
        """
        clear_values = ['' for i in range(14)]
        empty_rows = [clear_values for i in range(rows_cnt - 1)]
        self.sheet_obj.values().batchUpdate(
            spreadsheetId=sheet_id,
            body={"valueInputOption": "USER_ENTERED", 
                "data": [{"range": f"{list_title}!A2:N{rows_cnt}","majorDimension": "ROWS", "values": empty_rows}]
            }
        ).execute()
        print(f'Очистили лист {list_title}')

    def add_line(self, sheet_id, list_name, line_val):
        """ """
        err_cnt = 5
        err_msg = ''
        while err_cnt > 0:
            try:
                self.sheet_obj.values().append(
                    spreadsheetId=sheet_id, range=list_name, valueInputOption='USER_ENTERED', 
                    insertDataOption='INSERT_ROWS', body={"majorDimension": "ROWS","values": line_val}
                ).execute()
                print(f'Добавили строку на лист {list_name}')
                return
            except Exception as e:
                err_msg = str(e)
                print(f'{err_cnt}: {err_msg}')
                err_cnt -= 1
                time.sleep(60)
        print(f'Ошибка при отправке запроса add_line():\n{err_msg}', 'e')
        exit(1)

    def sort_lines(self, sheet_id, list_name, row_st=0, row_end=0, col_st=0, col_end=0, ind=0, ord='DESCENDING'):
        """ """
        list_id = self.sheet_list(sheet_id)[list_name]
        if row_end == 0:
            row_end = len(self.read_list(sheet_id, list_name)['values']) + 2
        if col_end == 0:
            col_end = 20
        sort_json_body = {"requests": [{"sortRange": {"range": { 
            "sheetId": list_id, "startRowIndex": row_st, "endRowIndex": row_end, "startColumnIndex": col_st, "endColumnIndex": col_end},
            "sortSpecs": [{"dimensionIndex": ind,"sortOrder": ord}]}}]
        }
        err_cnt = 5
        err_msg = ''
        while err_cnt > 0:
            try:
                self.sheet_obj.batchUpdate(spreadsheetId=sheet_id, body=sort_json_body).execute()
                print(f'Выполнили сортировку листа {list_name}')
                return
            except Exception as e:
                err_msg = str(e)
                print(f'{err_cnt}: {err_msg}')
                err_cnt -= 1
                time.sleep(60)
        print(f'Ошибка при отправке запроса sort_lines():\n{err_msg}', 'e')
        exit(1)

    def create_page(self, sheet_id, page_name):
        """
        Функция создания листа, если он еще не создан

        :param page_name: Название нового листа
        :param from_template: Шаблон листа для создания :class:`Worksheets <Worksheet>` (опционально)
        :param index: Индекс нового листа (опционально)
        :return: Экземпляр страницы документа
        """
        results = self.sheet_obj.batchUpdate(
            spreadsheetId = sheet_id,
            body = { "requests": [{"addSheet": {"properties": { "title": page_name,"gridProperties": {}}}}]
        }).execute()

        return results['replies'][0]['addSheet']['properties']

    def insert_range(self, sheet_id, list_name, data_list, row_st, row_end, col_st, col_end):
        """ """
        def get_leter(pos_val):    
            col_names = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
            if isinstance(pos_val, int):
                if pos_val < 26:
                    return col_names[pos_val]
                else:
                    return col_names[25-pos_val] + col_names[25-pos_val]
            return pos_val
        err_cnt = 5
        err_msg = ''
        while err_cnt > 0:
            try:
                self.sheet_obj.values().batchUpdate(spreadsheetId=sheet_id, 
                    body={"valueInputOption": "USER_ENTERED", "data": [
                        {"range": f"{list_name}!{get_leter(col_st)}{row_st + 1}:{get_leter(col_end)}{row_end + 1}",
                        "majorDimension": "ROWS", 
                        "values": data_list}]
                    }).execute()
                return
            except Exception as e:
                err_msg = str(e)
                print(f'{err_cnt}: {err_msg}')
                err_cnt -= 1
                time.sleep(60)
        print(f'Ошибка при отправке запроса insert_range():\n{err_msg}', 'e')
        exit(1)

    def set_cell_format(self, sheet_id, list_id, row_number=1, col_number=1,
                        bg_color=None, text_format=None, halignment=None):
        """
        Форматирование ячеек таблицы
        :param sheet_id: id листа
        :param row_index: Номер начальной строки
        :param row_number: Количество строк
        :param col_index: Номер начального столбца
        :param col_number: Количество столбцов
        :param bg_color: Цвет фона ячеек. Пример: {"red": 0.8, "green": 0.8, "blue": 0.8, "alpha": 1}
        :param text_format: Формат текста.  {"foregroundColor": {object (Color)},
                                             "foregroundColorStyle": {object (ColorStyle)},
                                             "fontFamily": string, "fontSize": integer, "bold": boolean,
                                             "italic": boolean, "strikethrough": boolean, "underline": boolean}
        :param halignment: Горизонтальное выравнивание текста. LEFT - Текст явно выравнивается по левому краю ячейки.
                                                             CENTER - Текст явно выравнивается по центру ячейки.
                                                              RIGHT - Текст явно выравнивается по правому краю ячейки.
        :return: Возвращает результат выполнения команды или None, если была ошибка
        """
        userEnteredFormat = {}
        json_body = {"requests": [{"repeatCell": {"range": {"sheetId": list_id,
                                                            "startRowIndex": row_number-1,
                                                            "endRowIndex": row_number,
                                                            "startColumnIndex": col_number-1,
                                                            "endColumnIndex": col_number},
                                                  "fields": "userEnteredFormat"}}]}
        if bg_color is not None:
            userEnteredFormat["backgroundColor"] = bg_color
        if text_format is not None:
            userEnteredFormat["textFormat"] = text_format
        if halignment is not None and halignment in ["LEFT", "CENTER", "RIGHT"]:
            userEnteredFormat["horizontalAlignment"] = halignment
        json_body["requests"][0]["repeatCell"]["cell"] = {"userEnteredFormat": userEnteredFormat}
        err_cnt = 5
        err_msg = ''
        while err_cnt > 0:
            try:
                self.sheet_obj.batchUpdate(spreadsheetId=sheet_id, body=json_body).execute()
                return
            except Exception as e:
                err_msg = str(e)
                print(f'{err_cnt}: {err_msg}')
                err_cnt -= 1
                time.sleep(60)
        print(f'Ошибка при отправке запроса insert_range():\n{err_msg}', 'e')
        exit(1)
