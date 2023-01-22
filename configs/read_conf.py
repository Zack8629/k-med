import json


def get_json_configs(file_conf: str):
    try:
        with open(file_conf, 'r') as conf_file:
            return json.loads(conf_file.read())

    except UnicodeDecodeError as e:
        return {}

    except Exception as exception:
        print(f'read_conf => {exception = }')
        return {}


app_configs_file = 'configs/app.set'
app_config_json = get_json_configs(app_configs_file)

parser_configs_file = 'configs/parser.set'
parser_config_json = get_json_configs(parser_configs_file)
