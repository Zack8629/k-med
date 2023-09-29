import sys

from parser import check_license_expiration_date, check_show_and_start, get_version


def start_window(App, Window, license_info):
    App = App
    Window = Window
    Window.show()

    license_type = None
    start_license = None
    stop_license = None
    last_run_date = None

    def license_failure(Window):
        Window.pars_button.setEnabled(False)
        Window.pars_button.setText('Срок действия лицензии истек!')
        Window.repaint()

    try:
        license_type = license_info[0][0]
        start_license = license_info[0][1]
        stop_license = license_info[0][2]
        last_run_date = license_info[0][3]

        if license_type == 'Trial':
            Window.setWindowTitle(f'Парсер v{get_version()} (Trial)')

        if not stop_license:
            try:
                stop_license = sys.argv[1]

            except IndexError:
                stop_license = ''

    except Exception as err:
        print(f'start_window => {err = }')
        license_failure(Window)

    if license_type:
        if not check_license_expiration_date(start_license, stop_license, last_run_date):
            license_failure(Window)

    if check_show_and_start(sys.argv[-1]):
        Window.the_button_was_clicked()
        sys.exit()

    App.exec()
