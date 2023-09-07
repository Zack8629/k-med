from PyQt6.QtWidgets import QDialog

from window.py import Ui_manual_window


class ManualWindow(QDialog, Ui_manual_window):
    def __init__(self, parent=None):
        QDialog.__init__(self, parent)
        self.setupUi(self)

        self.ok_btn.clicked.connect(self.close)
