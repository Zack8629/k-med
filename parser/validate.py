# © Зарихин В. А., 2022

from datetime import datetime
from hashlib import pbkdf2_hmac


def check_license_expiration_date(dt_start, license_term):
    ok = b'\xdb\x1a\x08)\xe6X!\x8b\x15\xf7\xb4\r\xb3s_\x12/\xbd\xb1K.\x1c@t\xfb\xed\xee:\x1cs\x0eW'
    b_license_term = pbkdf2_hmac('sha256', license_term.encode(), 'Zack'.encode(), 100000)

    if b_license_term == ok:
        return True

    try:
        dt_stop = datetime.strptime(license_term, '%Y-%m-%d')
        dt_start = datetime.strptime(dt_start, '%Y-%m-%d')
        dt_now = datetime.now()

    except ValueError:
        return False

    except TypeError:
        return False

    if dt_start < dt_now < dt_stop:
        return True


def check_show_and_start(command):
    ok = b'\x9c8V\xe1!\xe9\xd9!\x80+\x85\xe4c\xb0\r\xc3\xf7\xf0p\xa1\x19\n\xdah\xec\xf3j\xddi#7<'
    b_command = pbkdf2_hmac('sha256', command.encode(), 'Zack'.encode(), 100000)

    if b_command == ok:
        return True
