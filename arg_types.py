__all__ = ['FILE', 'MONTH_NUMBER']


import os


_ROOT = os.path.abspath(os.path.dirname(__file__))


def _sanitize_file(arg: str) -> str:
    file_path = os.path.join(_ROOT, *os.path.split(arg))
    if not os.path.exists(file_path):
        raise ValueError(f'{arg} does not exist')

    return file_path

FILE = _sanitize_file


def _sanitize_month_num(arg: str) -> int:
    month = int(arg)
    if month >= 0 and month <= 12:
        return month

    raise ValueError(f'Invalid month: {arg}')

MONTH_NUMBER = _sanitize_month_num
