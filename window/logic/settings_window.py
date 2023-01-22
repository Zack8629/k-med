# © Зарихин В. А., 2022

import json

from PyQt6.QtWidgets import QDialog

from configs import app_configs_file, app_config_json
from window.py import Ui_settings_window


class SettingsWindow(QDialog, Ui_settings_window):
    local_time_close = 3.2

    def __init__(self, parent=None):
        QDialog.__init__(self, parent)
        self.setupUi(self)

        app_config_json.setdefault('time_close', self.local_time_close)
        self.time_close_dbox.setValue(app_config_json['time_close'])

        self.cancel_btn.clicked.connect(self.close)
        self.apply_btn.clicked.connect(self.change_settings)

    def change_settings(self):
        app_config_json['time_close'] = round(self.time_close_dbox.value(), 1)

        try:
            with open(app_configs_file, 'w') as conf_file:
                conf_file.write(json.dumps(app_config_json))

        except Exception:
            self.local_time_close = app_config_json['time_close']

        self.close()
