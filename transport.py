#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
  Создание POST и GET запросов
  класс RPCCall - Авторизация на сайте, посылка POST и GET запросов, вывод результатов запроса в удобной форме.
  метод create_url - Получение url для запроса.
  метод get_file - Получение файла с url с помощью get запроса
  метод recordset_friendly - Преобразует словарь вида sbis-record (с секциями "d" и "s") в обычный json
  метод convert_recordset - Преобразование списка секции d из recorsetа в список словарей
  метод preparation_dict - Преобразование словаря. В словаре остаются/удаляются поля, указанные в списке fields
"""

import json
from copy import deepcopy
from requests import post, get



PATH_AUTH = "/auth/service/sbis-rpc-service300.dll"  # url для авторизации
BASE_URL = "cloud.sbis.ru"
HEADERS_DEFAULT = {'Content-type': 'application/json; charset=UTF-8',
                   'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; rv:18.0) Gecko/20100101 Firefox/18.0',
                   'Connection': 'keep-alive', }
SID_PRIVAT = "00000003-00000003-000a-0000000000000000"  # SID для приват логики
INFO_TO_CONSOLE = False  # Вывод в консоль url и json запроса

def get_google_token():
    token = get('http://10.76.120.189:8080/get_token').text
    return token
def recordset_friendly(recordset):
    """
    Метод преобразует словарь вида sbis-record (с секциями "d" и "s") в обычный json
    :param recordset: Словарь формата sbis-record
    :return: Преобразованный словарь
    """
    friendly = []
    if isinstance(recordset, dict) and 's' in recordset and 'd' in recordset:
        if all([(isinstance(el, list) and len(el) == len(recordset['s'])) for el in recordset['d']]):
            for elm in recordset['d']:
                friendly.append({key['n']: recordset_friendly(value) for key, value in zip(recordset['s'], elm)})
            return friendly
        friendly = {key['n']: recordset_friendly(value) for key, value in zip(recordset['s'], recordset['d'])}
        return friendly
    return recordset


def convert_recordset(section_d, section_s) -> list:
    """
    Преобразование списка секции d из recorsetа в список словарей
    :param section_d: секция "d" из recordseta
    :param section_s: секция "s" из recordseta
    :return: список словарей преобразованной секции "d"
    """
    if len(section_d) == 0 or len(section_s) == 0:
        return section_d
    section_s = [it["n"] for it in section_s]
    return [{section_s[num]: val for num, val in enumerate(it)} for it in section_d]


def preparation_dict(dct, fields, delete=False) -> dict:
    """
    Преобразование словаря. В словаре остаются только поля, указанные в списке fields, если delete=False.
    Или, поля из списка fields удаляются. Если при delete=False поля из списка fields не было в словаре dict, то оно
    добавляется со значением None.
    :param dct: Словарь
    :param fields: Список полей в формате list или str (разделитель ",")
    :param delete: False - остаются только указанные поля, True - указанные поля удаляются. По умолчанию - False.
    :return: Преобразованный словарь.
    """
    if isinstance(fields, str):
        fields = [elm.strip() for elm in fields.split(",")]
    if delete:
        return {key: val for key, val in dct.items() if key not in fields}
    return {key: dct[key] if key in dct else None for key in fields}


def create_url(site="test", base="cloud.sbis.ru", protocol="https") -> str:
    """
    Получение url для запроса
    :param site: Имя стенда
    :param base: Постоянная составляющая url.
    :param protocol: http или https
    :return: url в формате строки
    """
    return f"{protocol}://{str(site) + '-' if site and site not in ['prod', 'main'] else ''}{base}"


class RPCCall:
    """
    Выполнение POST и GET запросов
    """
    def __init__(self, protocol="https", site="test", base=BASE_URL, headers=None, sid=None,
                 login="viewer", password="Viewer1234", url=None):
        """
        :param protocol: Протокол передачи данных. По умолчанию https
        :param site: Имя хоста
        :param base: Основная часть url
        :param headers: Заголовки. По умолчинию: parametrs.client_headers
        :param site: Адрес сайта. Используется как замена параметру host, если адреса нет в parametrs.logs_login
        :param sid: SID. По умолчанию: None
        :param url: url пользователя в готовом виде. Отменяет все установки (protocol, site, base).
        """
        self.has_auth = False   # Проверка авторизации. True - авторизация пройдена.
        self.site = site
        self.sid = sid
        self.url = create_url(site=site, base=base, protocol=protocol) if url is None else url
        self.headers = deepcopy(HEADERS_DEFAULT) if headers is None else headers
        self.set_sid(sid)
        self.client = {"login": login, "password": password}
        self.response = None

    def set_sid(self, sid):
        """
        Задать SId для запросов
        :param sid: SID
        :return: нет
        """
        if sid is not None:
            self.sid = sid
            self.headers['X-SBISSessionID'] = self.sid
            self.has_auth = True

    def auth(self, login="", password="") -> bool:
        """
        Авторизация на сайте
        :param login: Логин
        :param password: Пароль
        :return: True - авторизация прошла успешно, False - ошибка авотризации
        """
        url = f'{self.url}{PATH_AUTH}'
        login = login if login else self.client["login"]
        password = password if password else self.client["password"]
        jsn = {"jsonrpc": "2.0", "protocol": 4, "method": "САП.АутентифицироватьРасш",
              "params": {"login": login, "password": password}, "id": 1}
        if INFO_TO_CONSOLE:
            print("-- Аутентификация --------")
            print(url)
            print(jsn)
        response = post(url=url, json=jsn, headers="")
        js_response = json.loads(response.text)
        if response.status_code == 200 and "error" not in js_response:
            # self.headers['Cookie'] = response.headers['set-cookie']
            self.set_sid(js_response['result']['d'][0][1])
            self.headers['X-SBISSessionID'] = json.loads(response.text)['result']['d'][0][1]
            self.sid = self.headers['X-SBISSessionID']
            if INFO_TO_CONSOLE:
                print(f"Аутентификация пройдена. SID - {self.sid}")
            self.has_auth = True
            return True
        print(f"Аутентификация не пройдена. status_code = {response.status_code}")
        print(f"Содержание ответа: {js_response}")
        self.has_auth = False
        return False

    def post_request(self, js=None, path='/service/sbis-rpc-service300.dll', url=""):
        """
        Запрос с помощью json
        :param js: json запроса
        :param path: путь запроса, по умолчанию: /service/sbis-rpc-service300.dll
        :param url: url запроса или полная строка запроса
        :return: Ответ запроса в зависимости от параметра reply_full
        """
        self.response = self.post_request_pass(js=js, path=path, url=url)
        return self

    def post_request_pass(self, js=None, path='/service/sbis-rpc-service300.dll', url=""):
        """
        POST Запрос с помощью json. В отличие от get_request возвращает результат запроса и не записывает его
        во внутреннюю переменную. Используется для создания многопоточности.
        :param js: json запроса
        :param path: путь запроса, по умолчанию: /service/sbis-rpc-service300.dll
        :param url: url запроса или полная строка запроса
        :return: Ответ запроса в зависимости от параметра reply_full
        """
        if js is None:
            print("Задайте json запроса.")
            return None
        url = f'{self.url}{path}' if url == "" else url
        if INFO_TO_CONSOLE:
            print(url)
            print(self.headers)
            print(js)
        response = post(url=url, json=js, headers=self.headers)
        if INFO_TO_CONSOLE:
            mess = "Запрос выполнен успешно." if response.status_code == 200 \
                else f"Ошибка выполнения запроса. {response.status_code}"
            print(mess)
        return response

    def get_request(self, url="", path="", params=""):
        """
        GET запрос
        :param url: url запроса или полная строка запроса
        :param path: часть url без хоста. Должна начинаться с '/'
        :param params: параметры к запросу, если нужно в виде строки. Начинаются с '?'
        :return: Ответ запроса в зависимости от параметра reply_full
        """
        self.response = self.get_request_pass(url=url, path=path, params=params)
        return self

    def get_request_pass(self, url="", path="", params=""):
        """
        GET Запрос. В отличие от get_request возвращает результат запроса и не записывает его во внутреннюю переменную.
        Используется для создания многопоточности.
        :param url: url запроса или полная строка запроса
        :param path: часть url без хоста. Должна начинаться с '/'
        :param params: параметры к запросу, если нужно в виде строки. Начинаются с '?'
        :return: Ответ запроса в зависимости от параметра reply_full
        """
        url = f"{url if url else self.site}{path}{params}"
        response = get(url, headers=self.headers)
        if INFO_TO_CONSOLE:
            mess = "Запрос выполнен успешно." if response.status_code == 200 \
                else f"Ошибка выполнения запроса. {response.status_code}"
            print(mess)
        return response

    def json(self, response=None) -> dict:
        """
        Выдаёт результат запроса в виде Json
        :param response: Ответ от запроса, передаваемый через параметр
        :return: response.json()
        """
        return self.response.json() if response is None else response.json()

    def recordset(self, response=None) -> dict:
        """
        Выдаёт результат запроса в виде Json
        :param response: Ответ от запроса, передаваемый через параметр
        :return: response.json()
        """
        return recordset_friendly(self.response.json()["result"] if response is None else self.json(response)["result"])

    def text(self, response=None) -> str:
        """
        Выдаёт результат запроса в виде текста
        :param response: Ответ от запроса, передаваемый через параметр
        :return: response.text в формате str
        """
        return self.response.text if response is None else response.text

    def raw(self):
        """
        Выдаёт сырой результат запроса
        :return: response
        """
        return self.response

    def code(self, response=None) -> int:
        """
        Выдаёт status_code запроса в формате int
        :param response: Ответ от запроса, передаваемый через параметр
        :return: status_code
        """
        return self.response.status_code if response is None else response.status_code

    def get_ver(self) -> str:
        """
        Получение версии сайта в виде строки.
        :return: Строка в формате <версия>-<бильд> от <дата>
        """
        def find_str(text_str, start_str, end_str):
            find_start = text_str.find(start_str) + len(start_str)
            find_end = text_str.find(end_str, find_start)
            return text_str[find_start:find_end]

        if self.has_auth:
            ver_text = self.get_request(f'{self.url}/ver.html').text()
            ver = find_str(ver_text, '(ver ', ') - ')
            build = find_str(ver_text, '/">', '</a>')
            date = find_str(ver_text, '</a> (', ') <br>').replace(' ', '')
            return f"{ver}-{build} от {date}"
        return ""


def get_file(url: str, name: str) -> bool:
    """ Получение файла с url с помощью get запроса
    :param url: Адрес файла в сети
    :param name: Имя выходного файла
    :return: True - удача или False - ошибка
    """
    try:
        rec = get(url, stream=True)
        with open(name, 'wb') as file:
            for chunk in rec.iter_content(chunk_size=1024):
                if chunk:
                    file.write(chunk)
    except Exception as exc:
        print(f"Ошибка скачивания файла: {exc}")
        return False
    return True


print(get_google_token())