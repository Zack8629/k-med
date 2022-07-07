from openpyxl import load_workbook
from pandas import read_excel

from a_scripts.services import print_data_for_line

result_file = 'готовый файл.xlsx'

ingosstrakh_file = 'список ингосстрах.XLS'
cogaz_file = 'список согаз.xls'
reso_file = 'список ресо.xls'
rosgosstrakh_file = 'список росгострах.xls'


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


def get_list_policies(writable_sheet, last_line_file, policies_column=10):
    start_line_file = 2

    list_policies = []
    for line_pol in range(start_line_file, last_line_file + 1):
        value_cell = writable_sheet.cell(row=line_pol, column=policies_column).value
        list_policies.append(value_cell)

    return list_policies


def save_file_to_exel(file_to_read, writable_file, file_to_write, sheet_num=0):
    writable_file.save(file_to_write)

    sheet_name = writable_file.sheetnames[sheet_num]

    data_frame = read_excel(file_to_write, sheet_name=sheet_num)
    data_frame.to_excel(file_to_write, sheet_name=sheet_name, encoding='utf-8', index=False)

    print(f'Data from "{file_to_read}" is written to "{file_to_write}"!')


def pars_ingosstrakh(file_to_read: str, file_to_write: str, sheet_num=0):
    def get_data_to_write():
        data_frame = read_excel(file_to_read, sheet_name=sheet_num)

        exclude_column = [8, 9, 13, 14, 15, 16, 17]

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

    def write_data():
        dict_to_write = {
            0: 10,
            1: 2,
            2: 3,
            3: 4,
            4: 6,
            5: 5,
            6: 22,
            7: 13,
            8: 7,
            9: 8,
            10: 23,
        }

        writable_file = load_workbook(file_to_write, read_only=False, keep_vba=True)
        writable_sheet = writable_file.worksheets[sheet_num]

        last_line_the_file = writable_sheet.max_row
        line_to_write = last_line_the_file + 1

        first_column = writable_sheet.min_column

        policies = get_list_policies(writable_sheet=writable_sheet,
                                     last_line_file=last_line_the_file,
                                     policies_column=10)

        data_to_write = get_data_to_write()

        policy_position = 0

        next_line = 0
        for line in data_to_write:
            if line[policy_position] in policies:
                print(f'The policy: "{line[policy_position]}" is already in the recorded file')
                continue

            writable_sheet.cell(row=line_to_write + next_line,
                                column=first_column).value = last_line_the_file + next_line

            for idx_value, value in enumerate(line):
                column_to_write = dict_to_write[idx_value]
                writable_sheet.cell(row=line_to_write + next_line,
                                    column=column_to_write).value = value

            next_line += 1

        save_file_to_exel(file_to_read, writable_file, file_to_write)

    write_data()


def pars_cogaz(file_to_read: str, file_to_write: str, sheet_num=0):
    def get_data_to_write():
        data_frame = read_excel(file_to_read, sheet_name=sheet_num)

        exclude_column = [10, 11]
        sep_column = [1, 2]

        start_line = 20
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

                elif num_column in sep_column:
                    valid_value = cell_value.title().split()
                    for el in valid_value:
                        data_line.append(el)
                    continue

                valid_value = determine_gender(cell_value)
                data_line.append(valid_value)

            list_data.append(data_line)

        return list_data

    def write_data():
        dict_to_write = {
            0: 2,
            1: 3,
            2: 4,
            3: 6,
            4: None,
            5: 5,
            6: 22,
            7: 10,
            8: 7,
            9: 8,
            10: 13,
        }

        writable_file = load_workbook(file_to_write, read_only=False, keep_vba=True)
        writable_sheet = writable_file.worksheets[sheet_num]

        last_line_the_file = writable_sheet.max_row
        line_to_write = last_line_the_file + 1

        first_column = writable_sheet.min_column

        policies = get_list_policies(writable_sheet=writable_sheet,
                                     last_line_file=last_line_the_file,
                                     policies_column=10)

        data_to_write = get_data_to_write()

        policy_position = 7

        next_line = 0
        for line in data_to_write:
            if line[policy_position] in policies:
                print(f'The policy: "{line[policy_position]}" is already in the recorded file')
                continue

            writable_sheet.cell(row=line_to_write + next_line,
                                column=first_column).value = last_line_the_file + next_line

            for idx_value, value in enumerate(line):
                column_to_write = dict_to_write[idx_value]
                if column_to_write:
                    writable_sheet.cell(row=line_to_write + next_line,
                                        column=column_to_write).value = value

            next_line += 1

        save_file_to_exel(file_to_read, writable_file, file_to_write)

    write_data()


def pars_reso(file_to_read: str, file_to_write: str, sheet_num=0):
    def get_data_to_write():
        data_frame = read_excel(file_to_read, sheet_name=sheet_num)

        exclude_column = [11, 12]
        sep_column = [2, 3, 8, 9]

        start_line = 7
        last_line = data_frame.shape[0]

        start_column = 2
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

                elif num_column in sep_column:
                    valid_value = cell_value.title().split()
                    for el in valid_value:
                        data_line.append(el)
                    continue

                valid_value = determine_gender(cell_value)
                data_line.append(valid_value)

            list_data.append(data_line)

        return list_data

    def write_data():
        dict_to_write = {
            0: 2,
            1: 3,
            2: 4,
            3: 6,
            4: None,
            5: 5,
            6: 22,
            7: 10,
            8: 7,
            9: None,
            10: 8,
            11: None,
            12: 13,
        }

        writable_file = load_workbook(file_to_write, read_only=False, keep_vba=True)
        writable_sheet = writable_file.worksheets[sheet_num]

        last_line_the_file = writable_sheet.max_row
        line_to_write = last_line_the_file + 1

        first_column = writable_sheet.min_column

        policies = get_list_policies(writable_sheet=writable_sheet,
                                     last_line_file=last_line_the_file,
                                     policies_column=10)
        print(policies)

        policy_position = 7

        next_line = 0
        for line in data_to_write:
            if line[policy_position] in policies:
                print(f'The policy: "{line[policy_position]}" is already in the recorded file')
                continue

            writable_sheet.cell(row=line_to_write + next_line,
                                column=first_column).value = last_line_the_file + next_line

            for idx_value, value in enumerate(line):
                column_to_write = dict_to_write[idx_value]
                if column_to_write:
                    writable_sheet.cell(row=line_to_write + next_line,
                                        column=column_to_write).value = value

            next_line += 1

        save_file_to_exel(file_to_read, writable_file, file_to_write)

    data_to_write = get_data_to_write()
    write_data()


def pars_rosgosstrakh(file_to_read: str, file_to_write: str, sheet_num=0):
    def get_data_to_write():
        data_frame = read_excel(file_to_read, sheet_name=sheet_num)

        exclude_column = []
        sep_column = []

        start_line = 6
        last_line = data_frame.shape[0]

        start_column = 2
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

                elif num_column in sep_column:
                    valid_value = cell_value.title().split()
                    for el in valid_value:
                        data_line.append(el)
                    continue

                valid_value = determine_gender(cell_value)
                data_line.append(valid_value)

            list_data.append(data_line)

        return list_data

    def write_data():
        dict_to_write = {
            0: 2,
            1: 3,
            2: 4,
            3: 6,
            4: None,
            5: 5,
            6: 22,
            7: 10,
            8: 7,
            9: None,
            10: 8,
            11: None,
            12: 13,
        }

        writable_file = load_workbook(file_to_write, read_only=False, keep_vba=True)
        writable_sheet = writable_file.worksheets[sheet_num]

        last_line_the_file = writable_sheet.max_row
        line_to_write = last_line_the_file + 1

        first_column = writable_sheet.min_column

        policies = get_list_policies(writable_sheet=writable_sheet,
                                     last_line_file=last_line_the_file,
                                     policies_column=10)
        print(policies)

        policy_position = 7

        next_line = 0
        for line in data_to_write:
            if line[policy_position] in policies:
                print(f'The policy: "{line[policy_position]}" is already in the recorded file')
                continue

            writable_sheet.cell(row=line_to_write + next_line,
                                column=first_column).value = last_line_the_file + next_line

            for idx_value, value in enumerate(line):
                column_to_write = dict_to_write[idx_value]
                if column_to_write:
                    writable_sheet.cell(row=line_to_write + next_line,
                                        column=column_to_write).value = value

            next_line += 1

        # save_file_to_exel(file_to_read, writable_file, file_to_write)

    data_to_write = get_data_to_write()
    print_data_for_line(data_to_write)
    # write_data()


def main():
    num_runs = 1
    print(f'num_runs = {num_runs}')

    for i in range(num_runs):
        print()
        # pars_ingosstrakh(ingosstrakh_file, result_file, sheet_num=0)
        print()
        # pars_cogaz(cogaz_file, result_file, sheet_num=0)
        print()
        # pars_reso(reso_file, result_file, sheet_num=0)
        print()
        pars_rosgosstrakh(rosgosstrakh_file, result_file, sheet_num=0)


if __name__ == '__main__':
    main()
