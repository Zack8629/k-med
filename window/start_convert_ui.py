import subprocess


def start_convert():
    list_ui = ['about_window',
               'easter_window',
               'main_window',
               'manual_window',
               'settings_window']

    for file in list_ui:
        subprocess.run(['pyuic6', '-o', f'py/{file}.py', f'ui/{file}.ui'])
        print(f'{file} done!')


if __name__ == '__main__':
    start_convert()
