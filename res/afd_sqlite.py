import sqlite3
from pprint import pprint as pp
from tabulate import tabulate
from res import afd_ggl

class Work_with_datebase():
    """ Работа с базой данных """

    def __init__(self):
        
        self.ggl = afd_ggl.GoogleSheet()
        database = 'database.db'
        self.con = sqlite3.connect(database)

    def create_table_v(self, table_name):
        """ Создать таблицу """
        # Сбор колонок и типов
        table = self.ggl.get_values(table_name)
        col_into, row_name = self.get_colums(table, table_name)
        # SQL
        cur = self.con.cursor()
        cur.execute(
            f"""DROP TABLE IF EXISTS {table_name};""")
        cur.execute(
            f"""CREATE TABLE IF NOT EXISTS {table_name} ({row_name})""")
        # pp(table[2:])
        cur.executemany(
            f"""INSERT INTO {table_name} VALUES ({col_into}) """, table)
        self.con.commit()
        self.con.close

    def create_table_aosr(self, table_name="АОСР"):
        """ Создать таблицу """
        # Сбор колонок и типов
        table = self.ggl.get_values(table_name)
        col_number = table[1].index('number')
        col_type = table[1].index('type')
        for n1 in range(len(table)):
            n2 = n1 - 1
            col_t = table[n1][col_type]
            col_n1 = table[n1][col_number]
            col_n2 = table[n2][col_number]
            if (col_t == 'АОСР' or col_t == 'ВК') and col_n1 == '':
                table[n1][col_number] = col_n2
        col_into, row_name = self.get_colums(table, table_name)
        # SQL
        cur = self.con.cursor()
        cur.execute(
            f"""DROP TABLE IF EXISTS {table_name};""")
        cur.execute(
            f"""CREATE TABLE IF NOT EXISTS {table_name} ({row_name})""")
        # pp(table[2:])
        cur.executemany(
            f"""INSERT INTO {table_name} VALUES ({col_into}) """, table)
        self.con.commit()
        self.con.close

    def create_table_g(self, table_name):
        """ Создать таблицу """
       # Сбор колонок и типов
        table = list(map(list, zip(*self.ggl.get_values(table_name))))
        col_into, row_name = self.get_colums(table, table_name)
        ### Работа с базой данных ###
        cur = self.con.cursor()
        cur.execute(
            f"""DROP TABLE IF EXISTS {table_name};""")
        cur.execute(
            f"""CREATE TABLE IF NOT EXISTS {table_name} ({row_name});""")
        cur.executemany(
            f"""INSERT INTO {table_name} VALUES ({col_into}) """, table)
        self.con.commit()
        self.con.close

    def drop_table(self, table_name):
        """ Удалить таблицу """
        cur = self.con.cursor()
        query = f'''DROP TABLE IF EXISTS {table_name}'''
        cur.execute(query)
        self.con.commit()
        self.con.close

    def insert_into(self, table_name, ggl_sheet):
        """ Добавить записи в таблицу """
        # Количество переменных
        col_name = ''
        count = 1
        for n in range(len(ggl_sheet[0])):
            if n == len(ggl_sheet[0]) - 1:
                col_name = col_name + '?'
            else:
                col_name = col_name + '?,'
        cur = self.con.cursor()
        # Larger example that inserts many records at a time
        ggl_table = ggl_sheet[3:]
        cur.executemany(
            f""" 
            INSERT INTO {table_name} 
            VALUES ({col_name}) """, ggl_table)
        self.con.commit()
        self.con.close

    def db_insert_into_row(self, table_name):
        """ Добавить запись в таблицу """
        cur = self.con.cursor()
        # Insert a row of data
        cur.execute(
            "INSERT INTO stocks VALUES ('2006-01-05','BUY','RHAT',100,35.14)")
        self.con.commit()
        self.con.close

    def db_insert_into_rows(self, table_name):
        """ Добавить запись в таблицу """
        cur = self.con.cursor()
        # Larger example that inserts many records at a time
        purchases = [
            # ('2006-03-28', 'BUY', 'IBM', 1000, 45.00),
            # ('2006-04-05', 'BUY', 'MSFT', 1000, 72.00),
            ('2006-04-09', 'SELL', 'IBM', 500, 53.00),
        ]
        cur.executemany(
            f""" 
            INSERT INTO {table_name} 
            VALUES (?,?,?,?,?) """, purchases)
        self.con.commit()
        self.con.close

    def db_read_one(self, table_name):
        """  """
        cur = self.con.cursor()
        t = ('RHAT',)
        cur.execute(f'''
        SELECT * FROM {table_name} WHERE symbol=?
        ''', t)
        print(cur.fetchone())

    def db_read_all(self, table_name):
        """  """
        cur = self.con.cursor()
        t = ('RHAT',)
        cur.execute(""" 
        SELECT * 
        FROM stocks 
        WHERE symbol=? 
        """, t)
        print(cur.fetchall())

    def db_read_iterator(self, query):
        """  """
        cur = self.con.cursor()
        for row in cur.execute(query):
            print(row)

    def db_query(self, query):
        cur = self.con.cursor()
        cur.execute(query)
        x = cur.fetchall()
        self.con.close
        return x

    def db_query_fetchone(self, query):
        cur = self.con.cursor()
        cur.execute(query)
        x = cur.fetchone()
        self.con.close
        return x

    def db_query_commit(self, query):
        cur = self.con.cursor()
        cur.execute(query)
        x = cur.fetchall()
        self.con.commit()
        self.con.close
        return x

    def testdb(self):
        text = 'текст'
        sheet_fio = ggl.get_values("ФИО")
        pp(sheet_fio[1])
        return text

    def get_colums(self, table, table_name,  row_tag=1):
        """ Определение кол-во и имена стол """
        row_col = table[row_tag]
        count_col = ''   
        name_col = ''   
        for n in range(len(row_col)):
            if n == len(row_col) - 1:
                count_col += '?'
                name_col += row_col[n] + ' text'
            else:
                count_col += '?,'
                name_col += row_col[n] + ' text, '
        return count_col, name_col



class Create_Datebase():
    """ Создание баз данных и заполнение """

    def __init__(self):
        """ Работа с базой данных """
        print(f"Create_Datebase.__init__")
        self.db = Work_with_datebase()

    def create_table_aosr(self, name_table='aosr', name_sheet='АОСР'):
        """ Таблица АОСР """
        print(f"Create_Datebase.create_table_aosr")
        ggl_aosr = self.ggl.get_values(name_sheet)
        # В таблице АОСР протягиваем номер акта по столбцу number
        col_number = ggl_aosr[0].index('number')
        col_type = ggl_aosr[0].index('type')
        length_sheet = len(ggl_aosr)
        for n1 in range(length_sheet):
            n2 = n1 - 1
            col_t = ggl_aosr[n1][col_type]
            col_n1 = ggl_aosr[n1][col_number]
            col_n2 = ggl_aosr[n2][col_number]
            if ((col_t == 'AOSR' or col_t == 'VK') and col_n1 == ''):
                ggl_aosr[n1][col_number] = col_n2
        self.db.create_table_v(name_table, ggl_aosr)

    def create_table_vk(self, name_table='vk', name_table_ggl='ВК'):
        """ Таблица ВК """
        print(f"Create_Datebase.create_table_vk")
        ggl_vk = self.ggl.get_values(name_table_ggl)
        self.db.create_table_v(name_table, ggl_vk)

    def create_table_fio(self, name_table='fio', name_table_ggl='ФИО'):
        """ Таблица ВК """
        print(f"Create_Datebase.create_table_fio")
        self.db.drop_table(name_table)
        ggl_fio = self.ggl.get_values(name_table_ggl)
        self.db.create_table_v(name_table, ggl_fio)

    def create_table_object(self, name_table='object', name_sheet='Объект'):
        """ Таблица Объект """
        print(f"Create_Datebase.create_table_object")
        ggl_object = self.ggl.get_values(name_sheet)
        self.db.create_table_g(name_table, ggl_object)

    def updb(self):
        """ Обновить базу данных """
        self.db.create_table_g('Объект')
        self.db.create_table_v('ФИО')
        self.db.create_table_aosr()
        self.db.create_table_v('ВК')

if __name__ == "__main__":
    x = Create_Datebase()
    x.updb()
    # x.create_table_object()
    # x.create_table_fio()
    # db = Work_with_datebase()
    
    
