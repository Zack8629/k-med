# © Зарихин В. А., 2022

from parser import Parser, get_version

parser_settings = {
    'ингосстрах': {
        'start': True,
        'dict_to_write': {
            'Номер полиса': 0,
            'Фамилия': 1,
            'Имя': 2,
            'Отчество': 3,
            'Дата рождения': 4,
            'Пол': 5,
            'Адрес проживания': 6,
            'Наименование программы': 7,
            'Дата прикрепления': 8,
            'Дата окончания': 9,
            'Расширение': 10,
            'Ограничение': 11,
            'СНИЛС': 12,
            'Место работы': 16,
        },
        'start_line_to_read': 12,
        'start_column_to_read': 1,
        'exclude_column': (8, 9, 13, 16, 17),
        'extra_cell': {
            '4 1': False,
        },

    },
    'согаз': {
        'start': True,
        'dict_to_write': {
            'Фамилия': 0,
            'Имя': 1,
            'Отчество': 2,
            'Дата рождения': 3,
            'Пол': 4,
            'Адрес проживания': 5,
            'Телефон пациента': 6,
            'Номер полиса': 7,
            'Дата прикрепления': 8,
            'Дата окончания': 9,
            'Наименование программы': 10,
            'Место работы': 11,
        },
        'start_line_to_read': 20,
        'start_column_to_read': 1,
        'exclude_column': [11],
        'sep_column': {
            1: None,
        },
    },
    'согаз Ф_И_О': {
        'start': True,
        'dict_to_write': {
            'Фамилия': 0,
            'Имя': 1,
            'Отчество': 2,
            'Дата рождения': 3,
            'Пол': 4,
            'Адрес проживания': 5,
            'Телефон пациента': 6,
            'Номер полиса': 9,
            'Дата прикрепления': 10,
            'Дата окончания': 11,
            'Наименование программы': 12,
            'Место работы': 13,
        },
        'start_line_to_read': 21,
        'start_column_to_read': 1,
        'exclude_column': [15],
    },
    'ресо': {
        'start': True,
        'dict_to_write': {
            'Фамилия': 0,
            'Имя': 1,
            'Отчество': 2,
            'Дата рождения': 3,
            'Пол': 4,
            'Адрес проживания': 5,
            'Телефон пациента': 6,
            'Номер полиса': 7,
            'Дата прикрепления': 8,
            'Дата окончания': 9,
            'Наименование программы': 10,
            'Место работы': 11,
        },
        'start_line_to_read': 7,
        'start_column_to_read': 2,
        'exclude_column': [12],
        'sep_column': {
            2: None,
        },
    },
    'росгострах': {
        'start': True,
        'dict_to_write': {
            'Фамилия': 0,
            'Имя': 1,
            'Отчество': 2,
            'Пол': 3,
            'Дата рождения': 4,
            'Адрес проживания': 5,
            'Телефон пациента': 6,
            'Номер полиса': 7,
            'Наименование программы': 10,
            'Дата прикрепления': 13,
            'Дата окончания': 14,
            'Место работы': 16,
        },
        'start_line_to_read': 6,
        'start_column_to_read': 2,
        'sep_column': {
            2: None,
        },
        'step_line': 3,
        'extra_cell': {
            '2 1': True,
        },
    },
    'альфа': {
        'start': True,
        'dict_to_write': {
            'Номер полиса': 0,
            'Фамилия': 1,
            'Имя': 2,
            'Отчество': 3,
            'Дата рождения': 4,
            'Адрес проживания': 5,
            'Место работы': 6,
            'Дата прикрепления': 7,
            'Дата окончания': 8,
            'Наименование программы': 9,
            'Пол': 10,
        },
        'start_line_to_read': 7,
        'start_column_to_read': 1,
        'sep_column': {
            2: None
        },
        'step_line': 9,
    },
    'ренессанс': {
        'start': True,
        'dict_to_write': {
            'Фамилия': 0,
            'Имя': 1,
            'Отчество': 2,
            'Дата рождения': 3,
            'Адрес проживания': 4,
            'Телефон пациента': 5,
            'Номер полиса': 6,
            'Наименование программы': 7,
            'Дата прикрепления': 9,
            'Дата окончания': 11,
            'Место работы': 13,
            'Пол': 14,
        },
        'start_line_to_read': 20,
        'start_column_to_read': 0,
        'exclude_column': [0, 3],
        'sep_column': {
            1: None,
        },
        'extra_cell': {
            '2 2': False,
            '4 2': True,
            '6 2': False,
        },
    },
    'согласие 13': {
        'start': True,
        'dict_to_write': {
            'Номер полиса': 0,
            'Фамилия': 1,
            'Имя': 2,
            'Отчество': 3,
            'Дата рождения': 4,
            'Адрес проживания': 5,
            'Телефон пациента': 6,
            'Место работы': 7,
            'Дата прикрепления': 8,
            'Дата окончания': 9,
            'Наименование программы': 10,
            'Пол': 11,
        },
        'start_line_to_read': 11,
        'start_column_to_read': 2,
        'exclude_column': [10],
        'sep_column': {
            3: None,
            5: '8-',
        },
        'step_line': 14,
    },
    'согласие 15': {
        'start': True,
        'dict_to_write': {
            'Номер полиса': 0,
            'Фамилия': 1,
            'Имя': 2,
            'Отчество': 3,
            'Дата рождения': 4,
            'Адрес проживания': 5,
            'Телефон пациента': 6,
            'Наименование программы': 8,
            'Место работы': 9,
            'Дата прикрепления': 10,
            'Дата окончания': 11,
            'Пол': 12,
        },
        'start_line_to_read': 13,
        'start_column_to_read': 2,
        'exclude_column': [10],
        'extra_cell': {
            '6 2': False,
            '8 3': False,
            '9 3': False,
            '9 5': False,
        },
    },
    'альянс': {
        'start': True,
        'dict_to_write': {
            'Номер полиса': 0,
            'Фамилия': 1,
            'Имя': 2,
            'Отчество': 3,
            'Дата рождения': 4,
            'Адрес проживания': 5,
            'Телефон пациента': 6,
            'Место работы': 7,
            'Дата прикрепления': 9,
            'Дата окончания': 10,
            'Наименование программы': 11,
            'Пол': 12,
        },
        'start_line_to_read': 16,
        'start_column_to_read': 1,
        'exclude_column': [1, 5, 8],
        'sep_column': {
            3: None,
        },
        'step_line': 14,
        'extra_cell': {
            '7 3': False,
            '5 3': True,
            '2 1': False,
        },
    },
}


def start_all_parse(show_policies=False, show_data=False, save=False, move_after_reading=False):
    from window import Window

    progress = 0
    step_progress = int(100 / len(parser_settings.keys()))

    print(f'Parser v{get_version()}')
    print(f'Start parsing')
    print()

    for oms in parser_settings:
        if parser_settings[oms]['start']:
            print(f'{oms = }')
            # for set_key, set_val in parser_settings[oms].items():
            #     print(f'{set_key} = {set_val}')

            Parser(folder_to_read=oms,
                   dict_to_write=parser_settings[oms]['dict_to_write'],
                   start_line_to_read=parser_settings[oms]['start_line_to_read'],
                   start_column_to_read=parser_settings[oms]['start_column_to_read'],

                   exclude_column=parser_settings[oms].get('exclude_column'),
                   sep_column=parser_settings[oms].get('sep_column'),
                   step_line=parser_settings[oms].get('step_line'),
                   extra_cell=parser_settings[oms].get('extra_cell'),

                   move_after_reading=move_after_reading,
                   show_policies=show_policies,
                   show_data=show_data,
                   save=save).pars()

            progress += step_progress
            Window.progress_bar.setProperty("value", progress)
            print(f'{progress = }%')

        print()

    print('Pars DONE!')
    print('Developed by Zarikhin')


if __name__ == '__main__':
    start_all_parse()
