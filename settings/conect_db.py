import os
import sqlite3
import sys


def resource_path(relative_path):
    # Получаем абсолютный путь к ресурсам.
    try:
        # PyInstaller создает временную папку в _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath('.')

    return os.path.join(base_path, relative_path)


def read_db(path, table):
    # Создание подключения к базе данных
    connect = sqlite3.connect(path)

    # Создание курсора
    cursor = connect.cursor()

    # Выполнение запроса на выборку данных
    cursor.execute(f'SELECT * FROM "{table}"')

    # Получение всех результатов
    results = cursor.fetchall()

    # Вывод результатов
    # for row_data in results:
    #     print(row_data)

    # Закрытие соединения
    connect.close()

    return results


def write_db(path, table, column, val):
    # Создание подключения к базе данных
    conn = sqlite3.connect(path)

    # Создание курсора
    cursor = conn.cursor()

    # Вставка данных в таблицу
    cursor.execute(f'UPDATE "{table}" SET "{column}" = "{val}"')

    # Сохранение изменений
    conn.commit()

    # Закрытие соединения
    conn.close()


if __name__ == '__main__':
    # rr = read_db('default_settings.db', 'default_settings')
    # rr = read_db('license.file', 'License_information')
    # print(f'{rr = }')

    # val = 'Test'
    # write_db('license.file', 'License_information', 'Last_run_date', val=val)

    # rr = read_db('default_settings.db', 'default_settings')
    # rr = read_db('license.file', 'License_information')
    # print(f'{rr = }')

    pass
