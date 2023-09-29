import json
import time

from PyQt6.QtWidgets import QMainWindow

from parser import get_version, start_all_parse
from settings import app_settings_json, app_settings_file
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

        # self.show_data.setEnabled(True)
        # self.show_data.addActions([self.close])

        # menubar -> menu
        self.settings.setEnabled(True)
        self.settings.triggered.connect(self.show_settings_window)
        self.exit.triggered.connect(self.close)

        # menubar -> help
        self.about.triggered.connect(self.show_about_window)
        self.manual.triggered.connect(self.show_manual_window)

        try:
            self.set_trigger_values()

        except KeyError as key_err:
            print(f'{key_err = }')
            print('save_current_trigger_values')
            self.save_current_trigger_values()

        except Exception as err:
            print(f'ParserWindow => {err = }')

    def show_settings_window(self):
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

    def set_trigger_values(self):
        self.move_after_reading.setChecked(app_settings_json['move_after_reading'])
        self.show_policies.setChecked(app_settings_json['show_policies'])
        self.show_data.setChecked(app_settings_json['show_data'])
        self.save.setChecked(app_settings_json['save'])
        self.close_after_done.setChecked(app_settings_json['close_after_done'])

    def exit_parser(self):
        self.save_current_trigger_values()
        self.close()

    def the_button_was_clicked(self):
        time_start = time.time()
        print(f'{time_start = }')

        self.move_after_reading.setEnabled(False)
        self.show_policies.setEnabled(False)
        self.show_data.setEnabled(False)
        self.save.setEnabled(False)
        self.close_after_done.setEnabled(False)

        self.pars_button.setEnabled(False)
        self.pars_button.setText('Парсинг...')
        self.repaint()

        self.save_current_trigger_values()

        self.pars()

        self.progress_bar.setProperty('value', 100)
        self.pars_button.setText('Готово!')
        self.repaint()

        time_stop = time.time()
        print(f'{time_stop = }')

        lead_time = time_stop - time_start
        print(f'{lead_time = }')

        if self.close_after_done.isChecked():
            time.sleep(app_settings_json['closing_time'])
            self.close()

    def save_current_trigger_values(self):
        app_settings_json['move_after_reading'] = self.move_after_reading.isChecked()
        app_settings_json['show_policies'] = self.show_policies.isChecked()
        app_settings_json['show_data'] = self.show_data.isChecked()
        app_settings_json['save'] = self.save.isChecked()
        app_settings_json['close_after_done'] = self.close_after_done.isChecked()

        try:
            with open(app_settings_file, 'w', encoding='utf-8') as conf_file:
                conf_file.write(json.dumps(app_settings_json, ensure_ascii=False))

        except Exception as err:
            print(f'ParserWindow => pars => {err = }')

    def pars(self):
        start_all_parse(move_after_reading=self.move_after_reading.isChecked(),
                        show_policies=self.show_policies.isChecked(),
                        show_data=self.show_data.isChecked(),
                        save=self.save.isChecked())
