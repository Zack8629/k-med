from datetime import datetime, timedelta
from hashlib import pbkdf2_hmac

from settings.conect_db import write_db


def check_license_expiration_date(start_license, stop_license, last_run_date):
    ok = b'\xdb\x1a\x08)\xe6X!\x8b\x15\xf7\xb4\r\xb3s_\x12/\xbd\xb1K.\x1c@t\xfb\xed\xee:\x1cs\x0eW'
    b_license_term = pbkdf2_hmac('sha256', stop_license.encode(), 'Zack'.encode(), 100000)

    if b_license_term == ok:
        return True

    try:
        dt_stop = datetime.strptime(stop_license, '%Y-%m-%d') + timedelta(days=1)
        dt_start = datetime.strptime(start_license, '%Y-%m-%d')
        dt_last_run = datetime.strptime(last_run_date, '%Y-%m-%d')
        dt_now = datetime.now()

        if dt_last_run > dt_now:
            return False

    except ValueError:
        return False

    except TypeError:
        return False

    except Exception as err:
        print(f'VALIDATE => {err = }')
        return False

    if dt_start < dt_now < dt_stop:
        days_left = dt_stop - dt_now
        print(f'Лицензии осталсоь {days_left}')

        write_db('settings/license.file', 'License_information', 'Last_run_date', dt_now.strftime('%Y-%m-%d'))
        return True


def check_show_and_start(command):
    ok = b'\x9c8V\xe1!\xe9\xd9!\x80+\x85\xe4c\xb0\r\xc3\xf7\xf0p\xa1\x19\n\xdah\xec\xf3j\xddi#7<'
    b_command = pbkdf2_hmac('sha256', command.encode(), 'Zack'.encode(), 100000)

    if b_command == ok:
        return True
