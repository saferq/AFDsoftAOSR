# -*- coding: utf-8 -*-
import os
import main

from PyQt5 import QtCore
from PyQt5 import QtWidgets
from PyQt5 import uic


class MyWindow(QtWidgets.QWidget):
    """ 
    Программа AFDsoftAOSR2019 v1.4

    Описание Программы

    """

    def __init__(self, parent=None):
        """ Конструтор """
        print('MyWindow.__init__')
        QtWidgets.QWidget.__init__(self, parent)
        Form, Base = uic.loadUiType("ui//formAOSR.ui")
        self.ui = Form()
        self.ui.setupUi(self)
        # Выбор рабочей директории
        self.folder_doc = ''
        if self.folder_doc == '':
            self.ui.lbDir.setText(os.path.abspath("") + "\\")
            self.folder_doc = os.path.abspath("") + "\\"
        # Назначение функций на кнопки
        self.ui.btnDir.clicked.connect(self.choose_dir)
        self.ui.btnACT.clicked.connect(self.creat_act)

    def creat_act(self):
        """ Сохранить АОСР """
        print('MyWindow.creat_act')
        self.act = main.CreatAct()
        txt_qlineAORS = self.ui.lineACT.text() # Текст в поле номера строк
        if len(txt_qlineAORS) != 0:
            row = self.act.list_numbers(txt_qlineAORS)
            self.act.creat_aosr(row, self.folder_doc)
        self.ui.lineACT.returnPressed.connect(self.ui.btnACT.click)

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
