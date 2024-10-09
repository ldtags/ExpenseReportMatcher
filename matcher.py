import os
import datetime as dt
import argparse as ap
import openpyxl as xl
from difflib import SequenceMatcher
from openpyxl.cell import Cell

from arg_types import MONTH_NUMBER


DRIVE = 'C:\\Users'
USER = 'ltags'
PATH = 'Future Energy Enterprises LLC\\FutEE Employee - Operations\\4. Monthly Expense Tool'
EXPENSE_TOOL_PATH = os.path.join(
    DRIVE,
    USER,
    *os.path.split(PATH)
)
MONTH_RB_LIMIT = 1
USAGE_STR = 'Usage: ./main -e [--employee] <str> [-m [--month] <int>]'
EXPENSE_REPORT_SHEET_NAME = 'Expense Report'
SIMILARITY_TOLERANCE = 0


def parse_args() -> ap.Namespace:
    parser = ap.ArgumentParser()

    parser.add_argument(
        '-e', '--employee',
        type=str,
        required=True
    )

    parser.add_argument(
        '-m', '--month',
        type=MONTH_NUMBER,
        default=dt.date.today().month
    )

    return parser.parse_args()


class Expense:
    def __init__(self, row: tuple[Cell, ...]):
        if len(row) < 11:
            raise RuntimeError(f'Row [{row}] is not properly formatted')

        self.transaction_date = row[0].value
        self.description = row[2].value
        self.category = row[3].value
        self.transaction_type = row[4].value
        self.memo = row[6].value
        self.expense_type = row[7].value
        self.client = row[8].value

        _post_date = row[1].value
        if _post_date is not None:
            self.post_date = _post_date
        else:
            self.post_date = None

        _amount = row[5].value
        try:
            self.amount = float(_amount)
        except ValueError as err:
            raise RuntimeError(f'Invalid amount: {_amount}') from err

        match row[9].value:
            case 'Billable':
                self.billable = True
            case 'Non-Billable':
                self.billable = False
            case other:
                raise RuntimeError(f'Invalid billable type: {other}')

        match row[10].value:
            case 'Reimbursable':
                self.reimbursable = True
            case 'Non-Reimbursable':
                self.reimbursable = False
            case other:
                raise RuntimeError(f'Invalid reimbursable type: {other}')

    def __eq__(self, o: object) -> bool:
        if not isinstance(o, Expense):
            return False

        return self.__dict__ == o.__dict__

    def __str__(self) -> str:
        _repr = '\n'
        _repr += '-' * 255
        _repr = '\n'
        _repr += ''.join((
            'Transaction'.ljust(25),
            '|',
            'Post Date'.ljust(25),
            '|',
            'Description'.ljust(30),
            '|',
            'Category'.ljust(25),
            '|',
            'Type'.ljust(8),
            '|',
            'Amount'.ljust(10),
            '|',
            'Memo'.ljust(50),
            '|',
            'Type'.ljust(18),
            '|',
            'Client'.ljust(10),
            '|',
            'Billable'.ljust(25),
            '|',
            'Reimbursable'.ljust(25),
            '\n'
        ))
        _repr += '-' * 255
        _repr += '\n'
        _repr += ''.join((
            f'{self.transaction_date}'.ljust(25),
            '|',
            f'{self.post_date or ""}'.ljust(25),
            '|',
            self.description.ljust(30),
            '|',
            self.category.ljust(25),
            '|',
            self.transaction_type.ljust(8),
            '|',
            f'{self.amount:.2f}'.ljust(10),
            '|',
            (self.memo or '').ljust(50),
            '|',
            self.expense_type.ljust(18),
            '|',
            self.client.ljust(10),
            '|',
            ('Billable' if self.billable else 'Non-Billable').ljust(25),
            '|',
            ('Reimbursable' if self.reimbursable else 'Non-Reimbursable').ljust(25),
            '\n'
        ))
        _repr += '-' * 255
        _repr += '\n'
        return _repr


def calculate_similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()


def get_expense_report_dir_path(month: int) -> str:
    """Returns the path to the folder containing all monthly expense reports.

    Raises a FileNotFoundError if no expense report folder is found.
    """

    month_name = dt.date(1900, month, 1).strftime('%b')
    dir_path = os.path.join(EXPENSE_TOOL_PATH, f'{month:02} {month_name}', 'Expense Reports')
    if not os.path.exists(dir_path):
        raise FileNotFoundError(f'No expense report directory found for {month_name}')

    return dir_path


def choose_file(options: list[str]) -> str | None:
    """Returns a user-selected file from the list of file names in `options`.

    Returns None if this operation is cancelled by the user.
    """

    print(f'Please select one of the expense reports found below [0 - {len(options) - 1}]:')
    for i, file in enumerate(options):
        print(f'{i} - {file}')

    selection = ''
    while selection == '':
        try:
            selection = input()
        except (EOFError, KeyboardInterrupt):
            return None

        try:
            return options[int(selection)]
        except (IndexError, ValueError):
            print('Please select a valid file number')
            selection = ''


def get_expense_reports(month: int, employee: str | None=None) -> list[str]:
    """Returns a list of the file names of all expense reports
    of the given month.

    If `employee` is included, results will be limited to that
    employee.
    """

    dir_path = get_expense_report_dir_path(month)
    all_expense_reports = [f for f in os.listdir(dir_path) if os.path.isfile(os.path.join(dir_path, f))]
    if employee is None:
        return all_expense_reports

    expense_reports: list[str] = []
    for expense_report in all_expense_reports:
        if employee in expense_report:
            expense_reports.append(expense_report)

    return expense_reports


def get_expense_report(month: int, employee: str) -> str | None:
    """Returns the expense report for the employee.

    If multiple expense reports are found, the user is prompted to
    select one.

    If no expense report is found or multiple are found but none are
    selected by the user, returns None.
    """

    month_name = dt.date(1900, month, 1).strftime('%b')
    expense_reports = get_expense_reports(month, employee)
    if expense_reports == []:
        raise FileNotFoundError(f'No expense reports found for {employee} in {month_name}')
    elif len(expense_reports) > 1:
        print(f'Multiple expense reports found for {employee} in {month_name}.')
        report_name = choose_file(expense_reports)
        if report_name is None:
            return None
    else:
        report_name = expense_reports[0]

    return report_name


def main():
    args = parse_args()
    employee = getattr(args, 'employee', None)
    if not isinstance(employee, str):
        raise RuntimeError(USAGE_STR)

    month = getattr(args, 'month', None)
    if not isinstance(month, int):
        raise RuntimeError(USAGE_STR)

    try:
        report_name = get_expense_report(month, employee)
    except FileNotFoundError as err:
        print(err.strerror)
        return

    if report_name is None:
        return

    dir_path = get_expense_report_dir_path(month)
    report_wb_path = os.path.join(dir_path, report_name)
    report_wb = xl.load_workbook(report_wb_path)
    report_sheet = report_wb[EXPENSE_REPORT_SHEET_NAME]
    expenses: dict[str, Expense | None] = {}

    # store all initial expenses into a dict
    for i, row in enumerate(report_sheet):
        if i < 5:
            continue

        if all([cell.value is None for cell in row[0:11]]):
            break

        expenses[f'{row[2].value};;;;{float(row[5].value):.2f}'] = None

    # compare previous expenses with current expenses to find matches
    for i in range(1, MONTH_RB_LIMIT + 1):
        rb_month = month - i
        rb_month_name = dt.date(1900, rb_month, 1).strftime('%b')
        try:
            rb_dir_path = get_expense_report_dir_path(rb_month)
        except FileNotFoundError as err:
            print(err.strerror)
            print(f'Skipping {rb_month_name}...')

        try:
            rb_report_name = get_expense_report(rb_month, employee)
        except FileNotFoundError as err:
            print(err.strerror)
            print(f'Skipping {rb_month_name}...')
            continue

        if rb_report_name is None:
            print(f'Skipping {rb_month_name}...')
            continue

        rb_report_wb_path = os.path.join(rb_dir_path, rb_report_name)
        rb_report_wb = xl.load_workbook(rb_report_wb_path)
        rb_report_sheet = rb_report_wb[EXPENSE_REPORT_SHEET_NAME]
        for i, row in enumerate(rb_report_sheet):
            if i < 5:
                continue

            if all([cell.value is None for cell in row[0:11]]):
                break

            rb_key = f'{row[2].value};;;;{float(row[5].value):.2f}'
            rb_expense = Expense(row)
            for key in expenses:
                description, amount = key.split(';;;;', 1)
                if float(amount) == rb_expense.amount:
                    similarity = calculate_similarity(description.lower(), rb_expense.description.lower())
                    if similarity >= SIMILARITY_TOLERANCE:
                        print(f'Possible match ({(similarity * 100):.4f}%):\n{description} (${amount}) {rb_expense}')

            try:
                expense = expenses[rb_key]
            except KeyError:
                continue

            if expense is None:
                expenses[rb_key] = rb_expense
            elif expenses[rb_key] != rb_expense:
                print(f'Expense conflict found at row {i} in {rb_report_name}')
                continue


if __name__ == '__main__':
    main()
