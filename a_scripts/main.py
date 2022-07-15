from datetime import datetime

from openpyxl import load_workbook
from pandas import read_excel

result_file = 'готовый файл.xlsx'

ingosstrakh_file = 'списки от СК/список ингосстрах.XLS'
cogaz_file = 'списки от СК/список согаз.xls'
reso_file = 'списки от СК/список ресо.xls'
rosgosstrakh_file = 'списки от СК/список росгострах.xls'
alpha_file = 'списки от СК/список Альфа страхование.xlsx'
renaissance_file = 'списки от СК/список ренессанс.xls'


class Parser:
    def __init__(self, file_to_read, file_to_write, sheet_num_to_read=0, sheet_num_to_write=0,
                 exclude_column=None, sep_column=None, start_line_to_read=0,
                 start_column_to_read=0, step_line=0, dict_to_write=None,
                 position_policy_in_data=0, show_policies=False, show_data=False, save=True):

        if sep_column is None:
            sep_column = []

        if exclude_column is None:
            exclude_column = []

        self.file_to_read = file_to_read
        self.sheet_num_to_read = sheet_num_to_read
        self.exclude_column = exclude_column
        self.sep_column = sep_column
        self.start_line_to_read = start_line_to_read
        self.start_column_to_read = start_column_to_read
        self.step_line = step_line
        self.show_data = show_data

        if dict_to_write is None:
            dict_to_write = {}

        self.file_to_write = file_to_write
        self.dict_to_write = dict_to_write
        self.position_policy_in_data = position_policy_in_data
        self.sheet_num_to_write = sheet_num_to_write
        self.show_policies = show_policies
        self.save = save

    def get_data_to_write(self):
        data_frame = read_excel(self.file_to_read, sheet_name=self.sheet_num_to_read)

        start_line = self.start_line_to_read
        last_line = data_frame.shape[0]

        start_column = self.start_column_to_read
        last_column = data_frame.shape[1]

        list_data = []
        next_line = start_line
        for num_line in range(start_line, last_line):

            val_line = str(data_frame.iloc[next_line, start_column])
            if val_line == 'nan' or val_line.isspace():
                next_line += self.step_line
                if next_line > last_line:
                    break

                val_line = str(data_frame.iloc[next_line, start_column])
                if val_line == 'nan' or val_line.isspace():
                    break

            data_line = []
            for num_column in range(start_column, last_column):
                cell_value = str(data_frame.iloc[next_line, num_column])

                try:
                    cell_value = datetime.strptime(cell_value,
                                                   '%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y')
                except ValueError:
                    pass

                if cell_value == 'nan' or cell_value.isspace() or num_column in self.exclude_column:
                    continue

                if num_column in self.sep_column:
                    values = cell_value.title().split()

                    if len(values) == 2:
                        values.append('')

                    for val in values:
                        data_line.append(val)
                    continue

                value = self.determine_gender(cell_value).title()
                data_line.append(value)

            list_data.append(data_line)
            next_line += 1

        return list_data

    def write_data(self, data_to_write):
        writable_file = load_workbook(self.file_to_write, read_only=False, keep_vba=True)
        writable_sheet = writable_file.worksheets[self.sheet_num_to_write]

        last_line_the_file = writable_sheet.max_row
        line_to_write = last_line_the_file + 1

        first_column = writable_sheet.min_column

        policies = self.get_list_policies(writable_sheet=writable_sheet,
                                          policies_column=10)

        if self.show_policies:
            print(policies)

        policy_position = self.position_policy_in_data

        next_line = 0
        for line in data_to_write:
            if line[policy_position] in policies:
                if self.show_policies:
                    print(f'The policy: "{line[policy_position]}" is already in the recorded file')
                continue

            writable_sheet.cell(row=line_to_write + next_line,
                                column=first_column).value = last_line_the_file + next_line

            for idx_value, value in enumerate(line):
                column_to_write = self.dict_to_write[idx_value]
                if column_to_write:
                    writable_sheet.cell(row=line_to_write + next_line,
                                        column=column_to_write).value = value

            next_line += 1

        if self.save:
            self.save_file_to_exel(writable_file)
        else:
            print(f'SAVE = {self.save}')

    @staticmethod
    def print_data_for_line(data):
        for li in data:
            print()

            for val in enumerate(li):
                print(val)

    @staticmethod
    def determine_gender(val):
        female_gender_list = ['WOMEN', 'ЖЕНСКИЙ', 'ЖЕН',
                              'Women', 'Женский', 'Жен', 'Ж',
                              'women', 'женский', 'жен', 'ж']

        male_gender_list = ['MEN', 'МУЖСКОЙ', 'МУЖ',
                            'Men', 'Мужской', 'Муж', 'М',
                            'men', 'мужской', 'муж', 'м']

        if val in female_gender_list:
            return female_gender_list[-1]
        elif val in male_gender_list:
            return male_gender_list[-1]
        else:
            return val

    @staticmethod
    def get_list_policies(writable_sheet, policies_column=10):
        start_line_file = 2
        last_line_file = writable_sheet.max_row

        list_policies = []
        for line_pol in range(start_line_file, last_line_file + 1):
            value_cell = writable_sheet.cell(row=line_pol, column=policies_column).value
            list_policies.append(str(value_cell))

        return list_policies

    def save_file_to_exel(self, writable_file):
        file_to_write = self.file_to_write
        sheet_num = self.sheet_num_to_write

        writable_file.save(file_to_write)

        sheet_name = writable_file.sheetnames[sheet_num]

        data_frame = read_excel(file_to_write, sheet_name=sheet_num)
        data_frame.to_excel(file_to_write, sheet_name=sheet_name, encoding='utf-8', index=False)

        print(f'Data from "{self.file_to_read}" is written to "{file_to_write}"!')

    def pars(self):
        try:
            data_to_write = self.get_data_to_write()

            if self.show_data:
                self.print_data_for_line(data_to_write)

            self.write_data(data_to_write=data_to_write)

        except FileNotFoundError as file_not_found:
            print(f'File not found! {file_not_found}')

        except KeyError as key_error:
            print(f'Key "{key_error}" not found!')


def ingosstrakh_pars(show_policies=False, show_data=False, save=False):
    exclude_column = [8, 9, 13, 14, 15, 16, 17]

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

    ingosstrakh_parser = Parser(file_to_read=ingosstrakh_file,
                                file_to_write=result_file,
                                exclude_column=exclude_column,
                                start_line_to_read=12,
                                start_column_to_read=1,
                                position_policy_in_data=0,
                                dict_to_write=dict_to_write,
                                show_policies=show_policies,
                                show_data=show_data,
                                save=save)

    ingosstrakh_parser.pars()


def cogaz_pars(show_policies=False, show_data=False, save=False):
    dict_to_write = {
        0: 2,
        1: 3,
        2: 4,
        3: 6,
        4: 5,
        5: 22,
        6: 10,
        7: 7,
        8: 8,
        9: 13,
    }

    cogaz_parser = Parser(file_to_read=cogaz_file,
                          file_to_write=result_file,
                          exclude_column=[10, 11],
                          sep_column=[1],
                          start_line_to_read=20,
                          start_column_to_read=1,
                          position_policy_in_data=6,
                          dict_to_write=dict_to_write,
                          show_policies=show_policies,
                          show_data=show_data,
                          save=save)

    cogaz_parser.pars()


def reso_pars(show_policies=False, show_data=False, save=False):
    dict_to_write = {
        0: 2,
        1: 3,
        2: 4,
        3: 6,
        4: 5,
        5: 22,
        6: 10,
        7: 7,
        8: 8,
        9: 13,
    }

    reso_parser = Parser(file_to_read=reso_file,
                         file_to_write=result_file,
                         exclude_column=[11, 12],
                         sep_column=[2],
                         start_line_to_read=7,
                         start_column_to_read=2,
                         position_policy_in_data=6,
                         dict_to_write=dict_to_write,
                         show_policies=show_policies,
                         show_data=show_data,
                         save=save)

    reso_parser.pars()


def rosgosstrakh_pars(show_policies=False, show_data=False, save=False):
    dict_to_write = {
        0: 2,
        1: 3,
        2: 4,
        3: 5,
        4: 6,
        5: 22,
        6: 10,
    }

    rosgosstrakh_parser = Parser(file_to_read=rosgosstrakh_file,
                                 file_to_write=result_file,
                                 sep_column=[2],
                                 start_line_to_read=6,
                                 start_column_to_read=2,
                                 step_line=3,
                                 position_policy_in_data=6,
                                 dict_to_write=dict_to_write,
                                 show_policies=show_policies,
                                 show_data=show_data,
                                 save=save)

    rosgosstrakh_parser.pars()


def alfa_pars(show_policies=False, show_data=False, save=False):
    dict_to_write = {
        0: 10,
        1: 2,
        2: 3,
        3: 4,
        4: 6,
        5: 22,
        6: 7,
        7: 8,
    }

    alfa_parser = Parser(file_to_read=alpha_file,
                         file_to_write=result_file,
                         exclude_column=[5, 8],
                         sep_column=[2],
                         start_line_to_read=7,
                         start_column_to_read=1,
                         step_line=9,
                         position_policy_in_data=0,
                         dict_to_write=dict_to_write,
                         show_policies=show_policies,
                         show_data=show_data,
                         save=save)

    alfa_parser.pars()


def renaissance_pars(show_policies=False, show_data=False, save=False):
    dict_to_write = {
        0: 2,
        1: 3,
        2: 4,
        3: 6,
        4: 22,
        5: 10,
    }

    renaissance_parser = Parser(file_to_read=renaissance_file,
                                file_to_write=result_file,
                                exclude_column=[0, 3],
                                sep_column=[1],
                                start_line_to_read=20,
                                start_column_to_read=0,
                                position_policy_in_data=5,
                                dict_to_write=dict_to_write,
                                show_policies=show_policies,
                                show_data=show_data,
                                save=save)

    renaissance_parser.pars()


def class_pars(show_policies=False, show_data=False, save=False):
    ingosstrakh_pars(show_policies=show_policies, show_data=show_data, save=save)
    cogaz_pars(show_policies=show_policies, show_data=show_data, save=save)
    reso_pars(show_policies=show_policies, show_data=show_data, save=save)
    rosgosstrakh_pars(show_policies=show_policies, show_data=show_data, save=save)
    alfa_pars(show_policies=show_policies, show_data=show_data, save=save)
    renaissance_pars(show_policies=show_policies, show_data=show_data, save=save)
    print('Pars DONE!')
    print()


def main():
    num_runs = 1
    print(f'num_runs = {num_runs}')
    print()

    for i in range(num_runs):
        class_pars(show_policies=False, show_data=False, save=True)


if __name__ == '__main__':
    main()
