import json
import os


def get_settings_as_json(settings_file: str):
    try:
        with open(settings_file, 'r', encoding='utf-8') as sett_file:
            return json.loads(sett_file.read())

    except UnicodeDecodeError:
        return {}

    except FileNotFoundError:
        print(f'Conf file not found!')
        print(f'Creating a file with default settings')

        try:
            with open(settings_file, 'w', encoding='utf-8') as sett_file:
                sett_file.write(json.dumps({}))

        except FileNotFoundError:
            os.makedirs(os.path.dirname(settings_file))

            with open(settings_file, 'w', encoding='utf-8') as sett_file:
                sett_file.write(json.dumps({}))

        except Exception as fail:
            print(f'Creating a file with default settings - FAIL!')
            print(f'{fail = }')

        return {}

    except Exception as exception:
        print(f'get_settings_as_json => {exception = }')

        return {}


app_settings_file = 'settings/app.set'
app_settings_json = get_settings_as_json(app_settings_file)

parser_settings_file = 'settings/parser.set'
parser_settings_json = get_settings_as_json(parser_settings_file)
