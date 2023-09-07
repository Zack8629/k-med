# © Зарихин В. А., 2022 - 2023

from settings.conect_db import read_db
from window import App, Window, start_window

if __name__ == '__main__':
    print('start app')

    cw = read_db('settings/default_settings.db', 'default_settings')
    print(f'{cw = }')

    start_window(App, Window, '2022-12-24', 'free')
