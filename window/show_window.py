# © Зарихин В. А., 2022

import sys
import time

from PyQt6.QtWidgets import QMainWindow, QDialog

from parser import (start_all_parse,
                    check_license_expiration_date,
                    check_show_and_start,
                    get_version, get_copyright_sign)
from window import Ui_Parser_Window, Ui_about_window, Ui_easter_window, Ui_settings_window


class ParserWindow(QMainWindow, Ui_Parser_Window):
    click_counter = 0

    def __init__(self, parent=None):
        QMainWindow.__init__(self, parent)
        self.setupUi(self)
        self.setWindowTitle(f'Парсер v{get_version()}')

        # self.settings.setEnabled(True)

        self.pars_button.clicked.connect(self.the_button_was_clicked)

        self.settings.triggered.connect(self.show_settings)
        self.exit.triggered.connect(self.close)

        self.about.triggered.connect(self.show_about_window)

    def show_settings(self):
        Settings = SettingsWindow(self)
        Settings.show()
        Settings.exec()

    def show_easter_window(self):
        Easter = EasterWindow(self)
        Easter.show()
        Easter.exec()

    def show_about_window(self):
        About = AboutWindow(self)
        About.show()
        About.exec()

    def the_button_was_clicked(self):
        self.move_after_reading.setEnabled(False)
        self.show_policies.setEnabled(False)
        self.show_data.setEnabled(False)
        self.save.setEnabled(False)
        self.close_after_done.setEnabled(False)

        self.pars_button.setEnabled(False)
        self.pars_button.setText('Парсинг...')
        self.repaint()

        self.pars()

        self.progress_bar.setProperty("value", 100)
        self.pars_button.setText('Готово!')
        self.repaint()

        time.sleep(2)
        self.close()

    def pars(self):
        start_all_parse(move_after_reading=self.move_after_reading.isChecked(),
                        show_policies=self.show_policies.isChecked(),
                        show_data=self.show_data.isChecked(),
                        save=self.save.isChecked())


class AboutWindow(QDialog, Ui_about_window):
    def __init__(self, parent=None):
        QDialog.__init__(self, parent)
        self.setupUi(self)

        self.setWindowTitle(f'О парсере v{get_version()}')
        self.text_varsion.setText(f'Парсер v{get_version()}')
        self.text_dev.setText(f'Разработал {get_copyright_sign()}')

        self.ok_btn.clicked.connect(self.close)


class EasterWindow(QDialog, Ui_easter_window):
    def __init__(self, parent=None):
        QDialog.__init__(self, parent)
        self.setupUi(self)


class SettingsWindow(QDialog, Ui_settings_window):
    def __init__(self, parent=None):
        QDialog.__init__(self, parent)
        self.setupUi(self)


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
        Window.pars_button.setText('Срок действия лицензии истек!')
        Window.repaint()

    if check_show_and_start(sys.argv[-1]):
        Window.the_button_was_clicked()
        sys.exit()

    App.exec()
