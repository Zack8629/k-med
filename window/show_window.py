# © Зарихин В. А., 2022

import sys

from parser import check_license_expiration_date, check_show_and_start


def start_window(App, Window, dt_start, license_term=''):
    if not license_term:
        try:
            license_term = sys.argv[1]

        except IndexError:
            license_term = ''

    App = App

    Window = Window
    Window.show()

    if not check_license_expiration_date(dt_start, license_term):
        Window.pars_button.setEnabled(False)
        Window.pars_button.setText('Срок действия лицензии истек!')
        Window.repaint()

    if check_show_and_start(sys.argv[-1]):
        Window.the_button_was_clicked()
        sys.exit()

    App.exec()
