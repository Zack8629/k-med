import subprocess


def start_convert():
    list_ui = ['about_window',
               'easter_window',
               'parser_window',
               'settings_window',
               'manual_window']

    try:
        for file in list_ui:
            subprocess.run(['pyuic6', '-o', f'py/{file}.py', f'ui/{file}.ui'])
            print(f'{file} done!')
    except Exception:
        pass


if __name__ == '__main__':
    start_convert()
