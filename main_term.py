import os
import re
import afd_pd


class Main():
    ''' Консольный вариант программы
        ----------------------------
    '''

    def __init__(self):
        os.system("mode con cols=80 lines=30")
        print("0 - для выхода")
        self.act = afd_pd.MainManage()

    def convet_text_in_numbers(self, text):
        """ Текст полный текст """
        # Начало
        # Замена символов
        text = re.sub(r' ', '', text)
        text = re.sub(r'[.]', ',', text)
        list_text = re.split(r',', text)
        list_numbers = []
        for n in list_text:
            if re.fullmatch(r'\d+', n) != None:
                list_numbers.append(int(n))
            elif re.fullmatch(r'\d+-\d+', n) != None:
                nn = re.split(r'-', n)
                nn = [int(x) for x in nn]
                n_min = min(nn)
                n_max = max(nn)
                n_list = list(range(n_min, n_max + 1))
                list_numbers = list_numbers + n_list
            else:
                pass
        numbers = sorted(set(list_numbers))
        return numbers

    def main(self):
        while True:
            print("""Ввести номера строк:""")
            a = input("")
            if a == '0':
                # os.system('cls')
                break
            list_rows = self.convet_text_in_numbers(a)
            self.act.create_aosr(list_rows)



if __name__ == '__main__':
    go = Main()
    go.main()
