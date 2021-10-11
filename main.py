from lib.amital import Amital
from lib.logic import Logic
from rich.console import Console
import time

if __name__ == '__main__':
    console = Console()
    app = Logic(console)
    app.validate_file()
    # If passed excel validation
    if app.is_valid():

        app.console.print('✅ Excel file is valid and can be processed')
        amital = Amital(console)
        amital.get_mouse_position()
        # amital.fill_row_in_journal_screen(app.valid_rows[0])

        amital.login_screen()
        amital.first_menu_screen()
        # start month filling
        time.sleep(3)
        for id, row in enumerate(app.valid_rows):
            if row[0] != app.month:
                # new month
                app.console.print(f'👁‍ New month {row[0]} detected')
                if app.month is not None:
                    amital.close_journal_screen()
                app.set_month(row[0])
                amital.new_journal_screen(app.month)

            amital.fill_row_in_journal_screen(row)
        amital.close_journal_screen()
    else:
        app.console.print('❌ Excel file is invalid and cannot be processed')