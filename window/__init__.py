# © Зарихин В. А., 2022

import sys

from PyQt6.QtWidgets import QApplication

from .logic.main_window import ParserWindow
from .show_window import start_window

App = QApplication(sys.argv)
Window = ParserWindow()
