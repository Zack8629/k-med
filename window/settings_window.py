# Form implementation generated from reading ui file 'ui/settings_window.ui'
#
# Created by: PyQt6 UI code generator 6.3.1
#
# WARNING: Any manual changes made to this file will be lost when pyuic6 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt6 import QtCore, QtGui, QtWidgets


class Ui_settings_window(object):
    def setupUi(self, settings_window):
        settings_window.setObjectName("settings_window")
        settings_window.setWindowModality(QtCore.Qt.WindowModality.ApplicationModal)
        settings_window.resize(333, 333)
        settings_window.setMinimumSize(QtCore.QSize(333, 333))
        settings_window.setMaximumSize(QtCore.QSize(333, 333))
        self.ok_btn = QtWidgets.QPushButton(settings_window)
        self.ok_btn.setGeometry(QtCore.QRect(50, 100, 89, 25))
        self.ok_btn.setObjectName("ok_btn")
        self.settings_text = QtWidgets.QLabel(settings_window)
        self.settings_text.setGeometry(QtCore.QRect(90, 40, 141, 20))
        self.settings_text.setObjectName("settings_text")
        self.cancel_btn = QtWidgets.QPushButton(settings_window)
        self.cancel_btn.setGeometry(QtCore.QRect(170, 100, 89, 25))
        self.cancel_btn.setObjectName("cancel_btn")

        self.retranslateUi(settings_window)
        QtCore.QMetaObject.connectSlotsByName(settings_window)

    def retranslateUi(self, settings_window):
        _translate = QtCore.QCoreApplication.translate
        settings_window.setWindowTitle(_translate("settings_window", "Настройки"))
        self.ok_btn.setText(_translate("settings_window", "Применить"))
        self.settings_text.setText(_translate("settings_window", "Окно настроек"))
        self.cancel_btn.setText(_translate("settings_window", "Отмена"))
