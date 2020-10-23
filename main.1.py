import pprint as pp
from res import afd_doc, afd_docxtpl, afd_ggl, afd_sqlite, afd_pd


class CreatAct():
    def __init__(self):
        self.docx = afd_doc.Filling_act()
        self.dctp = afd_docxtpl.DocxT()
        self.db = afd_sqlite.Work_with_datebase()


    def creat_aosr(self, list_row, path):
        print('CreatAct.creat_aosr')


    def creat_avk(self, row_text, path):
        print('CreatAct.creat_avk')
        list_acts = self.get_acts(row_text)
        for act in list_acts:
            self.dctp.vk_08_rd(path, act)


    def get_acts(self, row_text):
        print('CreatAct.get_act')
        # получить минимум максимум в числе
        row_text = row_text.replace(' ', '')
        if row_text.isdigit() == True:
            row1 = row_text
            row2 = row_text
        elif '-' in row_text:
            num_tire = row_text.split('-')
            row1 = min(num_tire)
            row2 = max(num_tire)
        # sql
        query = f""" 
        SELECT number FROM ВК 
        WHERE rowid >= {row1} AND rowid <= {row2};"""
        val = self.db.db_query(query)
        # список
        list_acts = []
        for i in val:
            for j in i:
                list_acts.append(j)
        return list_acts


    def list_numbers(self, txt_from_qline):
        """ Создание и заполнение списка строк """
        print('CreatAct.list_numbers')
        self.txt_from_qline = txt_from_qline.replace(
            ' ', '').replace('.', ',').split(",")
        # списки
        self.list_row = []
        num_buf = []
        for num in self.txt_from_qline:
            if num.isdigit() == True:
                self.list_row.append(int(num))
            elif '-' in num:
                num_tire = num.split('-')
                for num_form in num_tire:
                    if num_form.isdigit() == True:
                        num_buf.append(int(num_form))
                    else:
                        next
                num_min = min(num_buf)
                num_max = max(num_buf) + 1
                for val in range(num_min, num_max):
                    self.list_row.append(int(val))
                num_buf.clear()
        return self.list_row

    def list_acts_aosr(self, num_first, num_last):
        """  """
        query = f""" SELECT number FROM aosr
        LIMIT {num_last - num_first + 1} OFFSET {num_first - 4};"""
        txt1 = self.db.db_query(query)
        txt2 = []
        n2 = ''
        for n in txt1:
            if n[0] != '' and n[0] != n2:
                txt2.append(n[0])
                n2 = n[0]
        return txt2

    def create_database(self):
        """ Создание баз дынных """
        db = afd_sqlite.Create_Datebase()
        db.updb()

class MainManage():

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
        docx = afd_docxtpl.DocxT()
        docx.aosr_v2019("e:\\python\\AFDsoftAOSR\\", number, context, name_act)

    def create_aosr(self, row_first, row_last):
        """ Заполнение АОСР """
        # google
        ggl_ws = afd_ggl.GoogleSheet()
        sheet_object = list(map(list, zip(*ggl_ws.get_values("Объект"))))
        sheet_aosr = ggl_ws.get_values("АОСР")
        sheet_fio = ggl_ws.get_values("ФИО")
        # pandas
        pd = afd_pd.PandasWork()
        df_object = pd.create_df(sheet_object)
        df_aosr = pd.create_df(sheet_aosr, col_start_content=0)
        df_fio = pd.create_df(sheet_fio)
        # docx
        x = df_aosr[row_first:row_last + 1]
        x1 = x.number.unique()
        for number in x1:
            if number != '':
                print('--')
                fio_id = self.query_body(df_aosr, number, 'АОСР', 'fio', end='', end_last='')
                name_act = self.query_body(df_aosr, number, 'АОСР', 'name')            
                self.aosr_fill(df_object, df_aosr, df_fio, fio_id, number, name_act)
            else:
                pass
        

if __name__ in "__main__":
    """ Запуск кода """
    x = MainManage()
    # x.create_aosr('АР-12')
    n = 40
    n2 = 48
    x.create_aosr(n, n2)