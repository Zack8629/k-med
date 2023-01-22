# © Зарихин В. А., 2022

import json

from PyQt6.QtWidgets import QDialog

from window.py import Ui_settings_window


class SettingsWindow(QDialog, Ui_settings_window):
    json_config = {}

    def __init__(self, parent=None):
        QDialog.__init__(self, parent)
        self.setupUi(self)
        self.file_window_conf = parent.file_window_conf

        try:
            self.json_config = parent.get_json_config()

        except Exception:
            self.json_config.setdefault('window', {'time_close': 3.2})

        self.time_close_dbox.setValue(self.json_config['window']['time_close'])

        self.cancel_btn.clicked.connect(self.close)
        self.apply_btn.clicked.connect(self.change_settings)

    def change_settings(self):
        json_config_new = self.json_config
        json_config_new["window"]["time_close"] = round(self.time_close_dbox.value(), 1)

        with open(self.file_window_conf, 'w') as conf_file:
            conf_file.write(json.dumps(json_config_new))

        self.close()
