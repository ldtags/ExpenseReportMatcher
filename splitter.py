import os
import csv
import argparse as ap
import datetime as dt
from configparser import ConfigParser

from arg_types import FILE, MONTH_NUMBER


_ROOT = os.path.abspath(os.path.dirname(__file__))


class Expense:
    def __init__(self, row: tuple):
        self.transaction_date = dt.datetime.strptime(row['Transaction Date'], r'%m/%d/%Y')
        self.post_date = dt.datetime.strptime(row['Post Date'], r'%m/%d/%Y')
        self.description = row['Description']
        self.category = row['Category']
        self.type = row['Type']
        self.amount = row['Amount']
        self.memo = row['Memo']

    def as_dict(self) -> dict[str,]:
        d = {}
        for key, value in self.__dict__.items():
            if '_' in key:
                f_key = " ".join(w.capitalize() for w in key.split('_'))
            else:
                f_key = key.capitalize()
            d[f_key] = value

        return d

def parse_args() -> ap.Namespace:
    parser = ap.ArgumentParser()

    parser.add_argument(
        '-f', '--file',
        type=FILE,
        required=True,
        help='Relative path to the CSV report file'
    )

    parser.add_argument(
        '-m', '--month',
        type=MONTH_NUMBER,
        help='Expense report month, overrides current month',
        default=dt.date.today().month
    )

    return parser.parse_args()


def main():
    args = parse_args()
    file_path = getattr(args, 'file', None)
    if file_path is None:
        raise RuntimeError('Missing file path')

    month = getattr(args, 'month', None)
    if month is None:
        raise RuntimeError('Missing month')

    month_name = dt.date(1900, month, 1).strftime(r'%B')

    config_parser = ConfigParser()
    config_parser.read(os.path.join(_ROOT, 'config.ini'))
    num_map = dict(config_parser.items('card numbers'))

    expense_map: dict[str, list[Expense]] = {}
    for _, cardholder in num_map.items():
        expense_map[cardholder] = []

    with open(file_path, newline='') as fd:
        reader = csv.DictReader(fd)
        for row in reader:
            card_num = row['Card']
            try:
                expense_map[num_map[card_num]].append(Expense(row))
            except KeyError as err:
                raise RuntimeError(f'Unknown card number: {card_num}') from err

    field_names = ['Transaction Date', 'Post Date', 'Description', 'Category', 'Type', 'Amount', 'Memo']
    for cardholder, expenses in expense_map.items():
        file_name = f'{cardholder} {month_name} CSV.csv'
        with open(os.path.join(_ROOT, 'csv reports', file_name), 'w', newline='') as fd:
            writer = csv.DictWriter(fd, fieldnames=field_names)
            writer.writeheader()
            writer.writerows([expense.as_dict() for expense in expenses])


if __name__ == '__main__':
    main()
