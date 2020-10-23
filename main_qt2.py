# -*- coding: utf-8 -*-
import os
import main

from PyQt5 import QtCore
from PyQt5 import QtWidgets
from PyQt5 import uic


class MyWindow(QtWidgets.QWidget):
    """ 
    Программа AFDsoftAOSR2019 v202010-01
    Описание Программы

    """

    def __init__(self, parent=None):
        """ Конструтор """
        print('MyWindow.__init__')
        QtWidgets.QWidget.__init__(self, parent)
        Form, Base = uic.loadUiType("ui//formAOSR.ui")
        self.ui = Form()
        self.ui.setupUi(self)
        self.ready()

    def ready(self):
        """Класс TextSplitter используется для разбивки текста на слова

        Основное применение - парсинг логов на отдельные элементы
        по указанному разделителю.

        Note:
            Возможны проблемы с кодировкой в Windows
        
        Attributes
        ----------
        file_path : str
            полный путь до текстового файла
        lines : list
            список строк исходного файла

        Methods
        -------
        load()
            Читает файл и сохраняет его в виде списка строк в lines
        
        get_splitted(split_symbol=" ")
            Разделяет строки списка по указанному разделителю
            и возвращает результат в виде списка
        """
        # Выбор рабочей директории
        self.folder_doc = ''
        if self.folder_doc == '':
            self.ui.lbDir.setText(os.path.abspath("") + "\\")
            self.folder_doc = os.path.abspath("") + "\\"
        self.ui.btnDir.clicked.connect(self.choose_dir)
        # Кнопка для сохранения АКТ
        self.ui.btnAVK.clicked.connect(self.creat_avk)
        self.act = main.CreatAct()

    def creat_avk(self):
        """ Сохранить АВК """
        print('MyWindow.creat_avk')
        row_text = self.ui.lineAVK.text()
        self.act.creat_avk(row_text, self.folder_doc)
        self.ui.lineAVK.returnPressed.connect(self.ui.btnAVK.click)

    def creat_aosr(self):
        """ Сохранить АОСР """
        print('MyWindow.creat_aosr')
        txt_qlineAORS = self.ui.lineAOSR.text()
        if len(txt_qlineAORS) != 0:
            row = self.act.list_numbers(txt_qlineAORS)
            self.act.creat_aosr(row, self.folder_doc)
        self.ui.lineAOSR.returnPressed.connect(self.ui.btnAOSR.click)

    def choose_dir(self):
        """ выбор директории """
        print('MyWindow.choose_dir')
        folder_doc = str(QtWidgets.QFileDialog.getExistingDirectory(
            self, "Select Directory"))
        self.folder_doc = QtCore.QDir.toNativeSeparators(folder_doc) + "\\"
        if self.folder_doc != '':
            if len(self.folder_doc) >= 35:
                name_path = self.folder_doc[0:3].rstrip(
                ) + '..' + self.folder_doc[-30:].rstrip()
                self.ui.lbDir.setText(name_path)
            else:
                name_path = self.folder_doc.rstrip()
                self.ui.lbDir.setText(name_path)
        print(f'folder_doc: {self.folder_doc}')


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    window = MyWindow()
    window.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint)
    window.show()
    sys.exit(app.exec_())
