import json

from PyQt6.QtWidgets import QDialog

from settings import app_settings_file, app_settings_json
from window.py import Ui_settings_window


class SettingsWindow(QDialog, Ui_settings_window):
    default_closing_time = 3.2
    default_language = 'Русский'

    def __init__(self, parent=None):
        QDialog.__init__(self, parent)
        self.setupUi(self)

        app_settings_json.setdefault('current_language', self.default_language)
        # self.language_combo_box.setValue(app_settings_json['current_language'])

        app_settings_json.setdefault('closing_time', self.default_closing_time)
        self.time_close_dbox.setValue(app_settings_json['closing_time'])

        self.cancel_btn.clicked.connect(self.close)
        self.apply_btn.clicked.connect(self.change_settings)

    def change_settings(self):
        app_settings_json['current_language'] = self.default_language
        app_settings_json['closing_time'] = round(self.time_close_dbox.value(), 1)

        try:
            with open(app_settings_file, 'w', encoding='utf-8') as conf_file:
                conf_file.write(json.dumps(app_settings_json, ensure_ascii=False))

        except Exception as e:
            print(f'change_settings => {e = }')

        self.close()
