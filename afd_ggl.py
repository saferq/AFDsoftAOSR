# -*- coding: utf-8 -*-
import pprint as pp
import gspread
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials

from res import afd_config
# import afd_config

class GoogleSheet:
    """ Работа с google таблицей """
    # Проверка лицензии
    def check_license(self):
        """ Проверка лицензии """
     # Подключение
        print(f"Проверка лицензии:")
        id_l = afd_config.id_license
        scope = [
            'https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_dict(
            afd_config.cred_json,
            scope)
        client = gspread.authorize(creds)
        sh = client.open_by_key(afd_config.key_table_license)
        ws_check = sh.worksheet("check")
     # Получение данных таблицы
        date_sw = ws_check.get_all_values()[id_l][1]
        date_check = datetime.strptime(date_sw, "%d.%m.%Y")
        date_now = datetime.now()
     #  Проверка условия
        date_difference = date_check - date_now
        if date_difference.days >= 0:
            print(f"осталось {date_difference.days} дней до {date_sw}\n")
            return True
        else:
            print(f"лицензия до {date_sw} просим вас продлить\n")
            return False

    # Получить данные c листа таблицы
    def get_values(self, sheet_name):
        """ Получить данные таблицы """
        print(f"Получение данных с таблицы '{sheet_name}'")
     # Подключение к таблице
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_dict(
            afd_config.cred_json,
            scope)
        client = gspread.authorize(creds)
        sh = client.open_by_key(afd_config.key_table_work)
        ws_aosr = sh.worksheet(sheet_name)
        data_sheet = ws_aosr.get_all_values()
        return data_sheet

################ АОСР #################
    # Создает словарь для заполнения титульного листа
    def creat_dict_aosr_title(self, sheet_name='Объект'):
        """ 
        Получение информации c листа Объект\n
        Название объекта\n
        Данные организации 
        """
        # print("GoogleSheet.creat_dict_aosr_title")
        x1 = self.get_values(sheet_name)
        dict_aosr_1 = {}
        for n in range(2, len(x1)):
            dict_aosr_1[x1[n][1]] = x1[n][2:]
        return dict_aosr_1

    # Создает словарь из данных АОСР
    def creat_dict_aosr(self, d_aosr):
        """ Обработка данных с листа АОСР """
        # print("GoogleSheet.creat_dict_aosr")
        length_rows = len(d_aosr)
        column_number = d_aosr[0].index('number')
        # Списки и словари
        aosr_list = []  # список номеров актов
        aosr_list_n = []  # список первой и последний строки
        aosr_list_n1 = []  # список первой строки акта
        aosr_list_n2 = []  # список последние строки акта
        # Заполнение списка актов
        for row in range(3, length_rows):
            try:
                line_current = d_aosr[row][column_number]
                line_next = d_aosr[row + 1][column_number]
                if line_current != '':
                    aosr_list.append(d_aosr[row][column_number])
                    aosr_list_n1.append(row)
                elif line_current == '' and line_next != '':
                    aosr_list_n2.append(row)
                else:
                    pass
            except IndexError as identifier:
                aosr_list_n2.append(length_rows - 1)
            # Заполнение списка строк
        for n in range(0, len(aosr_list_n1)):
            aosr_list_n.append(
                list(range(aosr_list_n1[n], aosr_list_n2[n] + 1)))
        # Заполнение словаря
        dict_aosr = dict(zip(aosr_list, aosr_list_n))
        return dict_aosr
    # Создает список названий АОСР нужных для печати

    def create_list_need_act(self, d_aosr, list_num_row):
        """ Сборка списка актов необходимых для печати """
        # print("GoogleSheet.creat_dict_aosr")
        dict_aosr = self.creat_dict_aosr(d_aosr)
        list_acts = []
        for kye, rows in dict_aosr.items():
            if not set(list_num_row).isdisjoint(rows):
                list_acts.append(kye)
        return list_acts
    # Создает словарь для определенного акта

    def creat_dict_aosr_general(self, d_aosr, act):
        """ 
        Заполнение основного словаря \n
        d_aosr2 - Данные таблицы с АОСР \n
        act 	- Название акта \n
        """
        # print("GoogleSheet.creat_dict_aosr_general")
        dict_aosr = self.creat_dict_aosr(d_aosr)
        act_nd = dict.fromkeys(d_aosr[0], [])
        for n in dict_aosr[act]:
            if n == dict_aosr[act][0]:
                for n2 in d_aosr[0]:
                    n3 = d_aosr[0].index(n2)
                    act_nd[n2] = [d_aosr[n][n3]]
            elif n > dict_aosr[act][0]:
                if d_aosr[n][d_aosr[0].index("type")] == 'AOSR':
                    n_p3 = d_aosr[0].index("point_1_2")
                    act_nd['point_1_2'].append(d_aosr[n][n_p3])
                if d_aosr[n][d_aosr[0].index("type")] == 'VK':
                    n_p3 = d_aosr[0].index("point_1_2")
                    act_nd['point_3'].append(d_aosr[n][n_p3])
        return (act_nd)
    # Получаем данные ответственных лиц

    def get_values_aosr_fio(self, sheet_name='ФИО'):
        # print("GoogleSheet.get_values_aosr_fio")
        sheet_values = self.get_values(sheet_name)
        return sheet_values
    # Создаем словарь для опеределенного акта АОСР

    def creat_dict_aosr_fio(self, sheet_values, id):
        """ Словарь с фамилиями """
        # print("GoogleSheet.creat_dict_aosr_fio")
        number = sheet_values[0].index('id')
        a_type = sheet_values[0].index('type')
        description = sheet_values[0].index('description')
        col_1 = sheet_values[0].index('n')
        col_2 = sheet_values[0].index('position')
        col_3 = sheet_values[0].index('organization')
        col_4 = sheet_values[0].index('last_name')
        col_5 = sheet_values[0].index('initials')
        col_6 = sheet_values[0].index('order')
        col_7 = sheet_values[0].index('requisites')
        dict_fio = {}
        for row in sheet_values:
            id2 = row[number]
            if id == id2:
                x = row[a_type]
                try:
                    dict_fio[x].append(row)
                except KeyError:
                    dict_fio[x] = [row]

        return dict_fio

################ ВК ###################
    # Получаем данные ответственных лиц
    def get_values_vk_fio(self, sheet_name='ФИО'):
        # print("GoogleSheet.get_values_vk_fio")
        sheet_values = self.get_values(sheet_name)
        return sheet_values

    # Создаем словарь для опеределенного акта  VK
    def creat_dict_vk_fio(self, sheet_values, id1, sheet_name='ФИО'):
        """ Словарь с фамилиями """
        # print("GoogleSheet.creat_dict_vk_fio")
        number = sheet_values[0].index('id')
        a_type = sheet_values[0].index('type')
        description = sheet_values[0].index('description')
        col_1 = sheet_values[0].index('n')
        col_2 = sheet_values[0].index('position')
        col_3 = sheet_values[0].index('organization')
        col_4 = sheet_values[0].index('last_name')
        col_5 = sheet_values[0].index('initials')
        col_6 = sheet_values[0].index('order')
        col_7 = sheet_values[0].index('requisites')
        dict_fio = {}
        for row in sheet_values:
            id2 = row[number]
            if id1 == id2:
                x = row[a_type]
                try:
                    dict_fio[x].append(row)
                except KeyError:
                    dict_fio[x] = [row]

        return dict_fio

