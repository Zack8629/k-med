# © Зарихин В. А., 2022

from parser import Parser, get_version


def ingosstrakh_pars(show_policies=False, show_data=False, save=False, move_after_reading=False):
    exclude_column = (8, 9, 13, 16, 17)

    dict_to_write = {
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
    }

    extra_cell = {
        '4 1': False,
    }

    Parser(folder_to_read='ингосстрах',
           dict_to_write=dict_to_write,
           start_line_to_read=12,
           start_column_to_read=1,
           exclude_column=exclude_column,
           extra_cell=extra_cell,
           move_after_reading=move_after_reading,
           show_policies=show_policies,
           show_data=show_data,
           save=save).pars()


def cogaz_pars(show_policies=False, show_data=False, save=False, move_after_reading=False):
    dict_to_write = {
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
    }

    sep_column = {
        1: None,
    }

    Parser(folder_to_read='согаз',
           dict_to_write=dict_to_write,
           start_line_to_read=20,
           start_column_to_read=1,
           exclude_column=[11],
           sep_column=sep_column,
           move_after_reading=move_after_reading,
           show_policies=show_policies,
           show_data=show_data,
           save=save).pars()


def cogaz_f_i_o_pars(show_policies=False, show_data=False, save=False, move_after_reading=False):
    dict_to_write = {
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
    }

    Parser(folder_to_read='согаз Ф_И_О',
           dict_to_write=dict_to_write,
           start_line_to_read=21,
           start_column_to_read=1,
           exclude_column=[15],
           move_after_reading=move_after_reading,
           show_policies=show_policies,
           show_data=show_data,
           save=save).pars()


def reso_pars(show_policies=False, show_data=False, save=False, move_after_reading=False):
    dict_to_write = {
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
    }

    sep_column = {
        2: None,
    }

    Parser(folder_to_read='ресо',
           dict_to_write=dict_to_write,
           start_line_to_read=7,
           start_column_to_read=2,
           exclude_column=[12],
           sep_column=sep_column,
           move_after_reading=move_after_reading,
           show_policies=show_policies,
           show_data=show_data,
           save=save).pars()


def rosgosstrakh_pars(show_policies=False, show_data=False, save=False, move_after_reading=False):
    dict_to_write = {
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
    }

    sep_column = {
        2: None,
    }

    extra_cell = {
        '2 1': True,
    }

    Parser(folder_to_read='росгострах',
           dict_to_write=dict_to_write,
           start_line_to_read=6,
           start_column_to_read=2,
           sep_column=sep_column,
           step_line=3,
           extra_cell=extra_cell,
           move_after_reading=move_after_reading,
           show_policies=show_policies,
           show_data=show_data,
           save=save).pars()


def alfa_pars(show_policies=False, show_data=False, save=False, move_after_reading=False):
    dict_to_write = {
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
    }

    sep_column = {
        2: None
    }

    Parser(folder_to_read='альфа',
           dict_to_write=dict_to_write,
           start_line_to_read=7,
           start_column_to_read=1,
           sep_column=sep_column,
           step_line=9,
           move_after_reading=move_after_reading,
           show_policies=show_policies,
           show_data=show_data,
           save=save).pars()


def renaissance_pars(show_policies=False, show_data=False, save=False, move_after_reading=False):
    dict_to_write = {
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
    }

    sep_column = {
        1: None,
    }

    extra_cell = {
        '2 2': False,
        '4 2': True,
        '6 2': False,
    }

    Parser(folder_to_read='ренессанс',
           dict_to_write=dict_to_write,
           start_line_to_read=20,
           start_column_to_read=0,
           exclude_column=[0, 3],
           sep_column=sep_column,
           extra_cell=extra_cell,
           move_after_reading=move_after_reading,
           show_policies=show_policies,
           show_data=show_data,
           save=save).pars()


def consent_pars_13(show_policies=False, show_data=False, save=False, move_after_reading=False):
    dict_to_write = {
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
    }

    sep_column = {
        3: None,
        5: '8-',
    }

    Parser(folder_to_read='согласие 13',
           dict_to_write=dict_to_write,
           start_line_to_read=11,
           start_column_to_read=2,
           exclude_column=[10],
           sep_column=sep_column,
           step_line=14,
           move_after_reading=move_after_reading,
           show_policies=show_policies,
           show_data=show_data,
           save=save).pars()


def consent_pars_15(show_policies=False, show_data=False, save=False, move_after_reading=False):
    dict_to_write = {
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
    }

    extra_cell = {
        '6 2': False,
        '8 3': False,
        '9 3': False,
        '9 5': False,
    }

    Parser(folder_to_read='согласие 15',
           dict_to_write=dict_to_write,
           start_line_to_read=13,
           start_column_to_read=2,
           exclude_column=[10],
           extra_cell=extra_cell,
           move_after_reading=move_after_reading,
           show_policies=show_policies,
           show_data=show_data,
           save=save).pars()


def alliance_pars(show_policies=False, show_data=False, save=False, move_after_reading=False):
    dict_to_write = {
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
    }

    sep_column = {
        3: None,
    }

    extra_cell = {
        '7 3': False,
        '5 3': True,
        '2 1': False,
    }

    Parser(folder_to_read='альянс',
           dict_to_write=dict_to_write,
           start_line_to_read=16,
           start_column_to_read=1,
           exclude_column=[1, 5, 8],
           sep_column=sep_column,
           step_line=14,
           extra_cell=extra_cell,
           move_after_reading=move_after_reading,
           show_policies=show_policies,
           show_data=show_data,
           save=save).pars()


def start_all_parse(move_after_reading=False, show_policies=False, show_data=False, save=False):
    print(f'Parser v{get_version()}')
    print(f'Start parsing')
    print()

    ingosstrakh_pars(move_after_reading=move_after_reading,
                     show_policies=show_policies,
                     show_data=show_data,
                     save=save)

    cogaz_pars(move_after_reading=move_after_reading,
               show_policies=show_policies,
               show_data=show_data,
               save=save)

    cogaz_f_i_o_pars(move_after_reading=move_after_reading,
                     show_policies=show_policies,
                     show_data=show_data,
                     save=save)

    reso_pars(move_after_reading=move_after_reading,
              show_policies=show_policies,
              show_data=show_data,
              save=save)

    rosgosstrakh_pars(move_after_reading=move_after_reading,
                      show_policies=show_policies,
                      show_data=show_data,
                      save=save)

    alfa_pars(move_after_reading=move_after_reading,
              show_policies=show_policies,
              show_data=show_data,
              save=save)

    renaissance_pars(move_after_reading=move_after_reading,
                     show_policies=show_policies,
                     show_data=show_data,
                     save=save)

    consent_pars_13(move_after_reading=move_after_reading,
                    show_policies=show_policies,
                    show_data=show_data,
                    save=save)

    consent_pars_15(move_after_reading=move_after_reading,
                    show_policies=show_policies,
                    show_data=show_data,
                    save=save)

    alliance_pars(move_after_reading=move_after_reading,
                  show_policies=show_policies,
                  show_data=show_data,
                  save=save)

    print()
    print('Pars DONE!')
    print(f'Developed by Zarikhin')


if __name__ == '__main__':
    start_all_parse(False, False, False, False)
