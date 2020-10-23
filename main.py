import pprint as pp

from res import afd_doc, afd_docxtpl, afd_ggl, afd_sqlite


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

if __name__ in "__main__":

    x = CreatAct()
    x.create_database()
    # print(x.get_acts("10-15"))
