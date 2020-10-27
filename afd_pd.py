import pprint as pp
import pandas as pd
import os

import afd_ggl
import afd_docxtpl_for_pd

class MainManage():
    def create_df(self, date_df, col_name=1, col_start_content=2):
        """ Создание DataFrame """
        df = pd.DataFrame(date_df[col_start_content:], columns=date_df[col_name])
        return df
    
    def query_object(self, df, name_col):
        """ Сбор даных с листа Объект """
        text = ''
        for index, col in df.iterrows():
            text = text + col[name_col]
            return text

    def query_body(self, df, number, type_doc, column, end='\t ', end_last=''):
        """ Сбор даных с листа АОСР """
        text = ''
        df1 = df[(df.number == number) & (df.type == type_doc)]
        # print(df1)
        for index, col in df1.iterrows():
            text = text + col[column] + end
        text = text + end_last
        return text

    def query_fio(self, df, id1, id2, column):
        """ Сбор даных с листа fio """
        text = ''
        df1 = df[(df.id1 == id1) & (df.id2 == id2)]
        # print(df1)
        for index, col in df1.iterrows():
            text = text + col[column] + '\t '
        return text

    def aosr_fill(self, df_object, df_aosr, df_fio, fio_id, number, name_act):
        """ Заполнить АОСР """
        # docx
        context = {
            # Шапка
            "object": self.query_object(df_object, "object"),
            "custumer": self.query_object(df_object, "custumer"),
            "builder": self.query_object(df_object, "builder"),
            "project": self.query_object(df_object, "project"),
            "build_name": self.query_object(df_object, "build_name"),
            "number": number,
            "date1": self.query_body(df_aosr, number, 'АОСР', 'date2', end='', end_last=''),
            # Тело
            "point1": self.query_body(df_aosr, number, 'АОСР', 'name' ),
            "point2": self.query_body(df_aosr, number, 'АОСР', 'proj', end=' ', end_last='\t '),
            "point3": self.query_body(df_aosr, number, 'ВК', 'name', end=' ', end_last='\t '),
            "point4": self.query_body(df_aosr, number, 'ВК', 'proj', end=' ', end_last='\t '),
            "point5_1": self.query_body(df_aosr, number, 'АОСР', 'date2', end='', end_last=''),
            "point5_2": self.query_body(df_aosr, number, 'АОСР', 'date3', end='', end_last=''),
            "point6": self.query_body(df_aosr, number, 'АОСР', 'pril2', end=' ', end_last='\t '),
            "point7": self.query_body(df_aosr, number, 'АОСР', 'razr', end=' ', end_last='\t '),
            "point4_1": self.query_body(df_aosr, number, 'ВК', 'proj', end=' ', end_last='\t '),
            # ФИО
            "fio_ZK_f":  self.query_fio(df_fio, fio_id, 'ZK', 'fio_full'),
            "fio_ZK_s":  self.query_fio(df_fio, fio_id, 'ZK', 'fio'),
            "fio_GP_f":  self.query_fio(df_fio, fio_id, 'GP', 'fio_full'),
            "fio_GP_s":  self.query_fio(df_fio, fio_id, 'GP', 'fio'),
            "fio_SKK_f":  self.query_fio(df_fio, fio_id, 'SKK', 'fio_full'),
            "fio_SKK_s":  self.query_fio(df_fio, fio_id, 'SKK', 'fio'),
            "fio_AN_f":  self.query_fio(df_fio, fio_id, 'AN', 'fio_full'),
            "fio_AN_s":  self.query_fio(df_fio, fio_id, 'AN', 'fio'),
            "fio_PD_f":  self.query_fio(df_fio, fio_id, 'PD', 'fio_full'),
            "fio_PD_s":  self.query_fio(df_fio, fio_id, 'PD', 'fio'),
            "fio_SK_f":  self.query_fio(df_fio, fio_id, 'SK', 'fio_full'),
            "fio_SK_s":  self.query_fio(df_fio, fio_id, 'SK', 'fio'),
        }

        path = os.path.abspath(os.curdir).replace('\\', '\\\\') + '\\\\'
        # Создание 
        docx = afd_docxtpl_for_pd.DocxT()
        docx.aosr_v2019(path, number, context, name_act)

    def create_aosr(self, list_rows):
        """ Заполнение АОСР """
        # google
        ggl_ws = afd_ggl.GoogleSheet()
        sheet_object = list(map(list, zip(*ggl_ws.get_values("Объект"))))
        sheet_aosr = ggl_ws.get_values("АОСР")
        sheet_fio = ggl_ws.get_values("ФИО")
        # pandas
        df_object = self.create_df(sheet_object)
        df_aosr = self.create_df(sheet_aosr, col_start_content=0)
        df_fio = self.create_df(sheet_fio)
        # Собираем список актов
        rows_act = []
        for row in list_rows:
            if df_aosr.number[row - 1] not in rows_act:
                rows_act.append(df_aosr.number[row - 1])
        # docx
        print("Заполнение и сохранение актов: ")
        count = 0
        for number in rows_act:
            if number != '':
                fio_id = self.query_body(df_aosr, number, 'АОСР', 'fio', end='', end_last='')
                name_act = self.query_body(df_aosr, number, 'АОСР', 'name')            
                self.aosr_fill(df_object, df_aosr, df_fio, fio_id, number, name_act)
                count += 1
            else:
                pass
        print(f"Создано актов : {count}\n")

if __name__ in "__main__":
    """ Запуск кода """
    x = MainManage()
    # x.create_aosr('АР-12')
    n = 6
    n2 = 20
    x.create_aosr(n, n2)