import argparse as ap
import datetime as dt
from typing import Any

from src import split_expenses, match_expenses
from src.arg_types import MONTH_NUMBER, EXISTING_FILE, POSITIVE_INT


def parse_args() -> ap.Namespace:
    parser = ap.ArgumentParser()
    subparsers = parser.add_subparsers(help='Types of mode')

    match_parser = subparsers.add_parser('match')
    match_parser.add_argument(
        '-e', '--employee',
        type=str,
        required=True
    )

    match_parser.add_argument(
        '-m', '--month',
        type=MONTH_NUMBER,
        default=dt.date.today().month
    )

    match_parser.add_argument(
        '-r', '--runback',
        type=POSITIVE_INT,
        default=3
    )

    split_parser = subparsers.add_parser('split')
    split_parser.add_argument(
        '-f', '--file',
        type=EXISTING_FILE,
        help='Relative path to the CSV report file',
        required=True
    )

    split_parser.add_argument(
        '-m', '--month',
        type=MONTH_NUMBER,
        help='Expense report month, overrides current month',
        default=dt.date.today().month
    )

    return parser.parse_args()


def determine_mode(args: ap.Namespace) -> str:
    flags = list(args.__dict__.keys())
    if flags == ['employee', 'month', 'runback']:
        return 'match'

    if flags == ['file', 'month']:
        return 'split'

    raise RuntimeError('Could not determine the mode')


def get_arg(args: ap.Namespace, arg_name: str) -> Any:
    arg = getattr(args, arg_name, None)
    if arg is None:
        raise RuntimeError(f'Missing {arg_name}')

    return arg


if __name__ == '__main__':
    args = parse_args()
    match determine_mode(args):
        case 'match':
            employee = get_arg(args, 'employee')
            month = get_arg(args, 'month')
            runback = get_arg(args, 'runback')
            match_expenses(employee, month, runback)
        case 'split':
            file = get_arg(args, 'file')
            month = get_arg(args, 'month')
            split_expenses(file, month)
        case other:
            raise RuntimeError(f'Unknown mode: {other}')
