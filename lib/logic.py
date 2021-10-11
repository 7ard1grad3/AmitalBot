import glob
import os

import pandas as pd
from rich import pretty
from rich.table import Table

from config import *


class Logic:
    def __init__(self, console):
        pretty.install()
        self.console = console
        self.valid = True
        self.errors = []
        self.valid_rows = []
        self.month = None
        # init first steps
        self.file = Logic.check_excel_file()

    def validate_file(self, file=None):
        if file is None:
            file = self.file
        if file is not None:
            self.console.print(f"✅ Found excel file in [bold magenta]{file}[/bold magenta]")
        else:
            self.add_error(
                f"❌ Missing file. [bold magenta]Make sure to place file in {EXCEL_FOLDER} folder "
                f"with .xlsx format[/bold magenta]", "error")
        if self.is_valid():
            self.validate_and_clean_excel(file)

    def mark_as_invalid(self):
        self.valid = False

    def is_valid(self, print_table=True):
        if print_table:
            self.show_errors()
        return self.valid

    def show_errors(self):
        if len(self.errors) > 0:
            table = Table(show_header=True, header_style="bold magenta")
            table.add_column("Message", style="dim", width=60)
            table.add_column("Type", style="red")
            for error in self.errors:
                table.add_row(error['message'], error['type'])
            self.console.print(table)

    def add_error(self, message, type):
        type = type.upper()
        self.errors.append({
            "message": message,
            "type": type,
        })
        if type.upper() == 'ERROR':
            self.mark_as_invalid()

    @staticmethod
    def check_excel_file():
        try:
            os.chdir(f"./{EXCEL_FOLDER}")
            for file in glob.glob("*.xlsx"):
                return f"../excel/{file}"
        except FileNotFoundError:
            return None
        return None

    def validate_and_clean_excel(self, file):
        with self.console.status("[bold green]Working on tasks...") as status:
            try:
                df = pd.read_excel(file, sheet_name=0)
            except FileNotFoundError as e:
                self.console.log(f"file {e.filename} not found ", style="red on white")
                return False
            # check for required structure of the excel file
            required_fields = ['Filter month', 'Transaction Date', 'cart nr', 'cart', 'sum nis',
                               'sum usd', 'ref', 'shipment']
            for required_field in required_fields:
                if required_field not in df.columns:
                    self.add_error(f"Missing field {required_field}", 'error')
            data = pd.DataFrame(df, columns=required_fields)
            for id, row in enumerate(data.sort_values(by='Filter month').values):
                valid_row = []
                id += 1
                # month must be 1-12
                if type(row[0]) != int and 12 > row[0] > 1:
                    self.add_error(
                        f"'Filter month' must be number that lower then 12 and higher then 1. failed at row {id}",
                        'error')
                else:
                    row[0] = str(row[0])
                    if len(row[0]) == 1:
                        row[0] = "0"+row[0]
                    valid_row.append(row[0])
                # date
                date = str(row[1])
                if len(date) not in [7, 8]:
                    self.add_error(f"'Transaction Date' must be in length of 7 or 8 in format of [daymonthyear]. "
                                   f"failed at row {id}",
                                   'error')
                else:
                    if len(date) == 7:
                        date = "0" + date

                    day = int(date[:2])
                    # day column
                    valid_row.append(day)
                    # always 3 column
                    valid_row.append(3)
                    # transaction date
                    # TODO: need to remove rthis line
                    date = '30092021'
                    valid_row.append(date)

                row[2] = str(row[2])
                if len(row[2]) not in [7, 8]:
                    self.add_error(f"'cart nr' must be in length of 7 or 8. failed at row {id}", 'error')
                else:
                    if len(row[2]) == 7:
                        row[2] = "0" + row[2]
                    valid_row.append(row[2])

                row[3] = str(row[3])
                if len(row[3]) not in [7, 8]:
                    self.add_error(f"'cart' must be in length of 7 or 8. failed at row {id}", 'error')
                else:
                    if len(row[3]) == 7:
                        row[3] = "0" + row[3]
                    valid_row.append(row[3])

                # Add static value row
                valid_row.append("01")

                # sum nis
                if type(row[4]) != int and row[4] <= 0:
                    self.add_error(f"'sum nis' must be numeric and higher then 0. failed at row {id}", 'error')
                else:
                    valid_row.append(row[4])

                # sum usd
                if type(row[5]) != int and row[5] <= 0:
                    self.add_error(f"'sum usd' must be numeric and higher then 0. failed at row {id}", 'error')
                else:
                    valid_row.append(row[5])

                row[6] = str(row[6])
                if len(row[6]) > 8:
                    self.add_error(f"'ref' is higher then 8. Will be trimmed - at row {id}", 'warning')
                row[6] = row[6].split(' ')[0]
                valid_row.append(row[6][:8])

                row[7] = str(row[7])
                if len(row[7]) > 7:
                    self.add_error(f"'ref' is higher then 7. Will be trimmed - at row {id}", 'warning')
                valid_row.append(row[7][:7])

                self.valid_rows.append(valid_row)

    def set_month(self, month):
        self.month = month
