# © Зарихин В. А., 2022

import sys

from PyQt6.QtWidgets import QApplication

from .about_window import Ui_about_window
from .easter_window import Ui_easter_window
from .parser_window import Ui_Parser_Window
from .settings_window import Ui_settings_window
from .show_window import start_window, ParserWindow

App = QApplication(sys.argv)
Window = ParserWindow()
