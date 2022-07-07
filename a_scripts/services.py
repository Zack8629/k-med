import pandas as pd


def print_data_for_line(data):
    for li in data:
        print()

        for val in enumerate(li):
            print(val)


def _get_template_csv(file: str, path_csv_file: str, sheet_name=0):
    data_frame = pd.read_excel(file, sheet_name=sheet_name)
    name_file = file.split(sep='/')

    csv_file = f'{path_csv_file}/{name_file[-1]}.csv'
    data_frame.to_csv(csv_file)


if __name__ == '__main__':
    pass

    file = '../test_files/список росгострах.xls'
    path_csv_file = '../csv_files/'
    _get_template_csv(file, path_csv_file)
