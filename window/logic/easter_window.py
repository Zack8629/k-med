# © Зарихин В. А., 2022

from PyQt6.QtWidgets import QDialog

from window.py import Ui_easter_window


class EasterWindow(QDialog, Ui_easter_window):
    def __init__(self, parent=None):
        QDialog.__init__(self, parent)
        self.setupUi(self)

        self.ok_btn.clicked.connect(self.close)
