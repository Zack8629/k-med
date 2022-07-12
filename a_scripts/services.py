import pandas as pd


def print_data_for_line(data):
    for li in data:
        print()

        for val in enumerate(li):
            print(val)


def _get_template_csv(file_to_read, path_csv_file=None, sheet_name=0):
    if path_csv_file is None:
        path_csv_file = './csv_files/'

    data_frame = pd.read_excel(file_to_read, sheet_name=sheet_name)
    name_file = file_to_read.split(sep='/')

    name_csv_file = f'{path_csv_file}/{name_file[-1]}.csv'
    data_frame.to_csv(name_csv_file)


if __name__ == '__main__':
    pass

    file = '../test_files/списки от СК/список ренессанс.xls'
    path_csv_file = '../csv_files/'
    _get_template_csv(file, path_csv_file)
