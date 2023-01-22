# © Зарихин В. А., 2022

import json
import time

from PyQt6.QtWidgets import QMainWindow

from configs import app_config_json
from parser import get_version, start_all_parse
from window.logic.about_window import AboutWindow
from window.logic.easter_window import EasterWindow
from window.logic.manual_window import ManualWindow
from window.logic.settings_window import SettingsWindow
from window.py.main_window import Ui_main_window


class ParserWindow(QMainWindow, Ui_main_window):
    click_counter = 0

    def __init__(self, parent=None):
        QMainWindow.__init__(self, parent)
        self.setupUi(self)
        self.setWindowTitle(f'Парсер v{get_version()}')

        self.pars_button.clicked.connect(self.the_button_was_clicked)

        # menubar -> menu
        self.settings.setEnabled(True)
        self.settings.triggered.connect(self.show_settings)
        self.exit.triggered.connect(self.close)

        # menubar -> help
        self.about.triggered.connect(self.show_about_window)
        self.manual.triggered.connect(self.show_manual_window)

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

    def show_manual_window(self):
        Manual = ManualWindow(self)
        Manual.show()
        Manual.exec()

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

        self.progress_bar.setProperty('value', 100)
        self.pars_button.setText('Готово!')
        self.repaint()

        if self.close_after_done.isChecked():
            time.sleep(app_config_json['time_close'])
            self.close()

    def pars(self):
        start_all_parse(move_after_reading=self.move_after_reading.isChecked(),
                        show_policies=self.show_policies.isChecked(),
                        show_data=self.show_data.isChecked(),
                        save=self.save.isChecked())

