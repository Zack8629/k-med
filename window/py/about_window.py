# Form implementation generated from reading ui file 'ui/about_window.ui'
#
# Created by: PyQt6 UI code generator 6.4.0
#
# WARNING: Any manual changes made to this file will be lost when pyuic6 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt6 import QtCore, QtGui, QtWidgets


class Ui_about_window(object):
    def setupUi(self, about_window):
        about_window.setObjectName("about_window")
        about_window.setWindowModality(QtCore.Qt.WindowModality.ApplicationModal)
        about_window.resize(300, 150)
        about_window.setMinimumSize(QtCore.QSize(300, 150))
        about_window.setMaximumSize(QtCore.QSize(300, 150))
        self.ok_btn = QtWidgets.QPushButton(about_window)
        self.ok_btn.setGeometry(QtCore.QRect(100, 100, 89, 25))
        self.ok_btn.setObjectName("ok_btn")
        self.text_varsion = QtWidgets.QLabel(about_window)
        self.text_varsion.setGeometry(QtCore.QRect(90, 30, 181, 20))
        self.text_varsion.setObjectName("text_varsion")
        self.text_dev = QtWidgets.QLabel(about_window)
        self.text_dev.setGeometry(QtCore.QRect(30, 60, 251, 17))
        self.text_dev.setObjectName("text_dev")

        self.retranslateUi(about_window)
        QtCore.QMetaObject.connectSlotsByName(about_window)

    def retranslateUi(self, about_window):
        _translate = QtCore.QCoreApplication.translate
        about_window.setWindowTitle(_translate("about_window", "Dialog"))
        self.ok_btn.setText(_translate("about_window", "Ок"))
        self.text_varsion.setText(_translate("about_window", "Парсер v{get_version()}"))
        self.text_dev.setText(_translate("about_window", "Разработал  © Зарихин В. А., 2022"))
