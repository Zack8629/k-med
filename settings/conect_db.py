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
    connect = sqlite3.connect(resource_path(path))

    # Создание курсора
    cursor = connect.cursor()

    # Вставка данных в таблицу
    # cursor.execute("INSERT INTO mytable (name, age) VALUES (?, ?)", ('John', 25))

    # Выполнение запроса на выборку данных
    cursor.execute(f'SELECT * FROM {table}')

    # Получение всех результатов
    results = cursor.fetchall()

    # Вывод результатов
    for row_data in results:
        print(row_data)

    # Закрытие соединения
    connect.close()

    return results[0][1]


def write_db(path, table, column, val):
    # Создание подключения к базе данных
    conn = sqlite3.connect(resource_path(path))

    # Создание курсора
    cursor = conn.cursor()

    # Вставка данных в таблицу
    cursor.execute(f'UPDATE {table} SET {column} = ({val})')

    # Сохранение изменений
    conn.commit()

    # Закрытие соединения
    conn.close()


if __name__ == '__main__':
    rr = read_db('default_settings.db', 'default_settings')
    print(f'{rr = }')

    # write_db('db.sqlite', 'current_settings', 'closing_time', 5)

    rr = read_db('default_settings.db', 'default_settings')
    print(f'{rr = }')
