# © Зарихин В. А., 2022

import sys
import time

from PyQt6.QtCore import QSize, Qt
from PyQt6.QtWidgets import (QApplication, QWidget, QMainWindow, QPushButton, QVBoxLayout,
                             QLabel, QCheckBox)
from parser import get_version
from run_pars import start_all_parse


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle(f'Parser v{get_version()}')
        self.setFixedSize(QSize(327, 163))

        self.label = QLabel()

        self.button = QPushButton('Начать парсинг')
        self.button.clicked.connect(self.the_button_was_clicked)
        self.setCentralWidget(self.button)

        self.checkBox_move_after_reading = QCheckBox(self.centralWidget())
        self.checkBox_move_after_reading.setObjectName('checkBox_move_after_reading')
        self.checkBox_move_after_reading.setText('checkBox_move_after_reading')

        self.checkBox_show_policies = QCheckBox(self.centralWidget())
        self.checkBox_show_policies.setObjectName('checkBox_show_policies')
        self.checkBox_show_policies.setText('checkBox_show_policies')

        self.checkBox_show_data = QCheckBox(self.centralWidget())
        self.checkBox_show_data.setObjectName('checkBox_show_data')
        self.checkBox_show_data.setText('save')

        self.checkBox_save = QCheckBox(self.centralWidget())
        self.checkBox_save.setObjectName('checkBox_save')
        # self.formLayout.setWidget(3, QtWidgets.QFormLayout.ItemRole.FieldRole, self.checkBox_save)
        self.checkBox_save.setText('save')

        layout = QVBoxLayout()
        layout.addWidget(self.label)

    def the_button_was_clicked(self):
        self.button.setEnabled(False)
        self.button.setText('Parsing...')
        self.repaint()

        self.pars()

        self.button.setText('DONE!')
        self.repaint()

        time.sleep(1)
        self.close()

    @staticmethod
    def pars():
        start_all_parse(move_after_reading=False,
                        show_policies=False,
                        show_data=False,
                        save=False)


def start_window():
    app = QApplication(sys.argv)

    window = MainWindow()
    window.show()

    app.exec()
