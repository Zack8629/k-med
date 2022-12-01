# © Зарихин В. А., 2022

import sys
import time

from PyQt6 import uic
from PyQt6.QtWidgets import QApplication, QMainWindow

from parser.run_pars import start_all_parse
from parser.validate import check_license_expiration_date, check_show_and_start


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


def start_window(license_term=''):
    sys_argv = sys.argv

    if not license_term:
        try:
            license_term = sys_argv[1]

        except IndexError:
            pass

    app = QApplication(sys_argv)

    window = ParserWindow()
    window.show()

    if not check_license_expiration_date(license_term):
        window.pars_button.setEnabled(False)
        window.pars_button.setText('License is expired!')
        window.repaint()

    if check_show_and_start(sys_argv[-1]):
        window.pars()

    app.exec()
