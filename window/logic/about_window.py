# © Зарихин В. А., 2022

from PyQt6.QtWidgets import QDialog

from parser import get_version, get_copyright_sign
from window.py import Ui_about_window


class AboutWindow(QDialog, Ui_about_window):
    def __init__(self, parent=None):
        QDialog.__init__(self, parent)
        self.setupUi(self)

        self.setWindowTitle(f'О парсере v{get_version()}')
        self.text_varsion.setText(f'Парсер v{get_version()}')
        self.text_dev.setText(f'Разработал {get_copyright_sign()}')

        self.ok_btn.clicked.connect(self.close)
