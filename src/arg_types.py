__all__ = ['FILE', 'MONTH_NUMBER', 'POSITIVE_INT']


import os

from src import ROOT


def _sanitize_file(arg: str) -> str:
    file_path = os.path.join(ROOT, *os.path.split(arg))
    if not os.path.exists(file_path):
        raise ValueError(f'{arg} does not exist')

    return file_path

EXISTING_FILE = _sanitize_file


def _sanitize_month_num(arg: str) -> int:
    month = int(arg)
    if month >= 0 and month <= 12:
        return month

    raise ValueError(f'Invalid month: {arg}')

MONTH_NUMBER = _sanitize_month_num


def _sanitize_positive_int(arg: str) -> int:
    arg_int = int(arg)
    if arg_int < 1:
        raise ValueError(f'{arg} must be a positive integer')

    return arg_int

POSITIVE_INT = _sanitize_positive_int
