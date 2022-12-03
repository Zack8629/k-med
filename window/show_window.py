# © Зарихин В. А., 2022

import sys
import time

from PyQt6 import uic
from PyQt6.QtWidgets import QMainWindow

from parser import (start_all_parse,
                    check_license_expiration_date,
                    check_show_and_start,
                    get_version)


class ParserWindow(QMainWindow):
    pars_button = None
    progress_bar = None

    move_after_reading = None
    show_policies = None
    show_data = None
    save = None

    def __init__(self):
        QMainWindow.__init__(self)
        uic.loadUi('window/parser_window.ui', self)

        # self.progress_bar.hide()

        self.pars_button.clicked.connect(self.the_button_was_clicked)
        self.setWindowTitle(f'Парсер v{get_version()}')

    def the_button_was_clicked(self):
        self.pars_button.setEnabled(False)
        self.pars_button.setText('Parsing...')
        self.repaint()

        self.pars()

        self.progress_bar.setProperty("value", 100)
        self.pars_button.setText('DONE!')
        self.repaint()

        time.sleep(1.5)
        self.close()

    def pars(self):
        start_all_parse(move_after_reading=self.move_after_reading.isChecked(),
                        show_policies=self.show_policies.isChecked(),
                        show_data=self.show_data.isChecked(),
                        save=self.save.isChecked())


def start_window(App, Window, license_term=''):
    if not license_term:
        try:
            license_term = sys.argv[1]

        except IndexError:
            license_term = ''

    App = App

    Window = Window
    Window.show()

    if not check_license_expiration_date(license_term):
        Window.pars_button.setEnabled(False)
        Window.pars_button.setText('License is expired!')
        Window.repaint()

    if check_show_and_start(sys.argv[-1]):
        Window.the_button_was_clicked()
        sys.exit()

    App.exec()
