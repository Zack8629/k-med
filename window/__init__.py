import os
import sys

from PyQt6.QtWidgets import QApplication

from settings.conect_db import read_db
from .logic.main_window import ParserWindow
from .show_window import start_window

App = QApplication(sys.argv)
Window = ParserWindow()

try:
    license_file = 'settings/license.file'

    if os.path.exists(license_file):
        license_info = read_db(license_file, 'License_information')

    else:
        license_info = None

except Exception:
    license_info = None
