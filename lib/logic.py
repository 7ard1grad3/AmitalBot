import glob
import os
import re
from datetime import timedelta, datetime

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
            answer = input("Multiple months in file? Y/N").lower()
            if answer == 'y':
                multiple = True
            else:
                multiple = False
            self.validate_and_clean_excel(file, multiple)

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

    def validate_and_clean_excel(self, file, multiple_months=True):
        with self.console.status("[bold green]Working on tasks...") as status:
            try:
                df = pd.read_excel(file, sheet_name=0)
            except FileNotFoundError as e:
                self.console.log(f"file {e.filename} not found ", style="red on white")
                return False
            # check for required structure of the excel file
            required_fields = ['FILTER MONTH', 'Transaction Date', 'CART NR', 'Column1', 'USD',
                               'USD2', 'Number', 'Shipment Name']
            for required_field in required_fields:
                if required_field not in df.columns:
                    self.add_error(f"Missing field {required_field}", 'error')
            data = pd.DataFrame(df, columns=required_fields)
            if multiple_months:
                data = data.sort_values(by='Transaction Date').values
            else:
                # Don't sort data if not multiple months
                data = data.values

            for id, row in enumerate(data):
                valid_row = []
                id += 1
                # month must be 1-12
                row[0] = int(row[0])
                if type(row[0]) != int and 12 > row[0] > 1:
                    self.add_error(
                        f"'FILTER MONTH' must be number that lower then 12 and higher then 1. failed at row {id}",
                        'error')
                else:
                    row[0] = str(row[0])
                    if len(row[0]) == 1:
                        row[0] = "0" + row[0]
                    valid_row.append(row[0])
                # date
                date = row[1]
                if type(date) is not pd._libs.tslibs.timestamps.Timestamp:
                    date = datetime.strptime(date, '%m/%d/%Y')
                    if type(date) is not pd._libs.tslibs.timestamps.Timestamp:
                        self.add_error(f"'Transaction Date' must be datetime in format of month/day/year. "
                                       f"failed at row {id} - value {date}",
                                       'error')
                else:
                    # check if saturday
                    if date.day_of_week == 5:
                        if date.day >= 28:
                            # self.add_error(f"{date.day} of {date.month} is Saturday. day will be reduced", 'warning')
                            date = date - timedelta(days=1)
                        else:
                            # self.add_error(f"{date.day} of {date.month} is Saturday. day will be added", 'warning')
                            date = date + timedelta(days=1)

                    date_string = date.strftime('%d%m%Y')
                    day = int(date.day)
                    # day column
                    valid_row.append(day)
                    # always 3 column
                    valid_row.append(3)
                    # transaction date
                    valid_row.append(date_string)

                row[2] = str(row[2]).replace("-", "").replace("/", "")
                if row[2] == "02003699":
                    row[2] = "02003698"
                if len(row[2]) not in [7, 8]:
                    self.add_error(f"'CART NR' must be in length of 7 or 8. failed at row {id}", 'error')
                else:
                    if len(row[2]) == 7:
                        row[2] = "0" + row[2]
                    valid_row.append(row[2])

                row[3] = str(row[3]).replace("-", "").replace("/", "")
                if len(row[3]) not in [7, 8]:
                    self.add_error(f"'Column1' must be in length of 7 or 8. failed at row {id} value: {row[3]}",
                                   'error')
                else:
                    if len(row[3]) == 7:
                        row[3] = "0" + row[3]
                    valid_row.append(row[3])

                # Add static value row
                valid_row.append("01")

                # sum nis
                if type(row[4]) != int and type(row[4]) != float:
                    self.add_error(f"'USD' must be numeric. failed at row {id}", 'error')
                else:
                    valid_row.append(row[4])

                # sum usd
                if type(row[5]) != int and type(row[5]) != float:
                    self.add_error(f"'USD2' must be numeric. failed at row {id}", 'error')
                else:
                    valid_row.append(row[5])

                row[6] = str(row[6])
                # if len(row[6]) > 8:
                #     self.add_error(f"'Number' is higher then 8. Will be trimmed - at row {id}", 'warning')

                row[6] = re.sub('[^A-Za-z0-9]+', '', row[6].split(' ')[0])
                if row[6] == "" or row[7] is None:
                    row[6] = "N/A"
                valid_row.append(row[6][:7])

                row[7] = re.sub('[^A-Za-z0-9]+', '', str(row[7]))
                if row[7] == "" or row[7] is None:
                    row[7] = "N/A"
                    # if len(row[7]) > 7:
                #     self.add_error(f"'Shipment Name' is higher then 7. Will be trimmed - at row {id}", 'warning')
                valid_row.append(row[7][:8])

                self.valid_rows.append(valid_row)

    def set_month(self, month):
        self.month = month
