import os

import openpyxl
import pandas as pd

from a_scripts.services import print_data_for_line, _get_template_csv

cwd = f'{os.getcwd()}/..'

dirname_source_file = f'{cwd}/test_files/файлы/списки от СК/'
source_files = os.listdir(dirname_source_file)
list_source_files = list(map(lambda name: os.path.join(dirname_source_file, name), source_files))
ing_file_test = list_source_files[-2]

result_file_test = f'{cwd}/test_files/Пример (загруженный).xlsx'

source_file_test = '../test_files/файлы/списки от СК/список ингосстрах.XLS'

source_file = '../working_files/файлы/списки от СК/список ингосстрах.XLS'
result_file = f'{cwd}/working_files/Файл для заполнения.xlsx'

path_to_csv_file = f'{cwd}/csv_files/'


def determine_gender(_val):
    title_val = _val.title()

    female_gender_list = ['WOMEN', 'ЖЕНСКИЙ', 'ЖЕН',
                          'Women', 'Женский', 'Жен', 'Ж',
                          'women', 'женский', 'жен', 'ж']

    male_gender_list = ['MEN', 'МУЖСКОЙ', 'МУЖ',
                        'Men', 'Мужской', 'Муж', 'М',
                        'men', 'мужской', 'муж', 'м']

    if title_val in female_gender_list:
        return female_gender_list[-1]
    elif title_val in male_gender_list:
        return male_gender_list[-1]
    else:
        return title_val


def get_data(file: str, sheet_name=0):
    data_frame = pd.read_excel(file, sheet_name=sheet_name)

    exclude_column = [8, 9, 10, 13, 14, 15, 16, 17]

    start_line = 12
    last_line = data_frame.shape[0]

    start_column = 1
    last_column = data_frame.shape[1]

    list_data = []
    for num_line in range(start_line, last_line):
        val_line = str(data_frame.iloc[num_line, start_column])
        if val_line == 'nan' or val_line.isspace():
            break

        data_line = []
        for num_column in range(start_column, last_column):
            cell_value = str(data_frame.iloc[num_line, num_column])

            if cell_value == 'nan' or cell_value.isspace() or num_column in exclude_column:
                continue

            valid_value = determine_gender(cell_value)
            data_line.append(valid_value)

        list_data.append(data_line)

    return list_data


def write_data_to_file(file_to_write: str, data, sheet_to_write=0):
    dict_to_write = {
        0: 10,
        1: 2,
        2: 3,
        3: 4,
        4: 6,
        5: 5,
        6: 22,
        7: 7,
        8: 8,
        9: 23,
    }

    writable_file = openpyxl.load_workbook(file_to_write, read_only=False, keep_vba=True)
    writable_sheet = writable_file.worksheets[sheet_to_write]

    last_line_the_file = writable_sheet.max_row
    line_to_write = last_line_the_file + 2

    first_column = writable_sheet.min_column

    last_serial_number = last_line_the_file - 1

    for idx_line, line in enumerate(data):
        writable_sheet.cell(row=line_to_write + idx_line,
                            column=first_column).value = last_serial_number + idx_line

        for idx_value, value in enumerate(line):
            writable_sheet.cell(row=line_to_write + idx_line,
                                column=dict_to_write[idx_value]).value = value

    writable_file.save(file_to_write)

    name_file = file_to_write.split(sep='/')
    print(f'Write DATA to {name_file[-1]} DONE!')


def main():
    data = get_data(ing_file_test)
    write_data_to_file(result_file_test, data)


if __name__ == '__main__':
    main()
