# © Зарихин В. А., 2022

import sys

from PyQt6.QtWidgets import QApplication

from .show_window import start_window, ParserWindow

App = QApplication(sys.argv)
Window = ParserWindow()
