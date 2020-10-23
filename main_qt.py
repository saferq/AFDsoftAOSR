# -*- coding: utf-8 -*-
import os
from PyQt5 import QtCore
from PyQt5 import QtWidgets

from ui import Ui_formACT
import main



class MyWindow(QtWidgets.QWidget):
    """ Программа ИСП """
    def __init__(self, parent=None):
        """ Конструтор """
        print('MyWindow.__init__')
        QtWidgets.QWidget.__init__(self, parent)
        self.ui = Ui_formACT.Ui_Form()
        self.ui.setupUi(self)
        self.ready()

    def ready(self):
        # Выбор рабочей директории
        self.folder_doc = ''
        if self.folder_doc == '':
            self.ui.lbDir.setText(os.path.abspath("") + "\\")
            self.folder_doc = os.path.abspath("") + "\\"
        self.ui.btnDir.clicked.connect(self.choose_dir)
        # Кнопка для обновления 
        self.ui.db_update.clicked.connect(self.db_update)
        # Кнопка для сохранения АВК
        self.ui.btnAVK.clicked.connect(self.creat_avk)
        self.act = main.CreatAct()

    def db_update(self):
        """ Обновить данные """
        print('MyWindow.db_update')
        self.act.create_database()
        self.ui.lineAVK.returnPressed.connect(self.ui.db_update.click)

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
