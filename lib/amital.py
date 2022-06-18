import time

import keyboard
import pyautogui
import pyperclip as pc
from pynput.mouse import Controller

# import clipboard as c
import config
from config import *


class Amital:
    def __init__(self, console):
        self.console = console
        self.window_position = None
        self.mouse = Controller()

    def get_mouse_position(self):
        self.console.print(f'⬜️Open Amital and place the mouse in the middle of Amital screen '
                           f'then press "{config.START_KEY}" on the keyboard')
        while True:  # Key listener to start the process
            try:  # used try so that if user pressed other than the given key error will not be shown
                if keyboard.is_pressed(config.START_KEY):
                    print(f'☑️Amital screen detected at ({self.mouse.position})')
                    self.window_position = self.mouse.position
                    time.sleep(0.02)
                    break  # finishing the loop
            except:
                break  # if user pressed a key other than the given key the loop will break

        time.sleep(2)

    def focus_on_screen(self, click=True):
        if click:
            pyautogui.click(self.window_position)
        else:
            pyautogui.moveTo(self.window_position)

    def login_screen(self):
        self.focus_on_screen()
        time.sleep(0.02)
        pyautogui.write(AMITAL_USER)
        time.sleep(0.02)
        pyautogui.press('enter')
        pyautogui.write(AMITAL_PASSWORD)
        pyautogui.press('enter')
        time.sleep(0.02)
        self.console.print('✅ Logging in to first menu')
        time.sleep(0.5)

    def first_menu_screen(self):
        self.focus_on_screen()
        time.sleep(0.02)
        pyautogui.write("H2")
        time.sleep(0.02)
        pyautogui.press('enter')
        time.sleep(0.02)
        self.console.print('✅ Going to H2')
        time.sleep(0.5)

    def new_journal_screen(self, moth):
        self.focus_on_screen()
        time.sleep(0.02)
        pyautogui.write("1")
        time.sleep(0.02)
        pyautogui.press('enter')
        time.sleep(0.02)

        pyautogui.write(moth)
        pyautogui.write(".")
        pyautogui.write("2021")
        pyautogui.press('enter')
        pyautogui.write("01")
        time.sleep(1)
        pyautogui.press('enter')
        pyautogui.press('enter')
        time.sleep(0.02)
        self.console.print(f'✅ {moth} month selected')
        time.sleep(0.5)

    def close_journal_screen(self):
        self.focus_on_screen()
        time.sleep(0.02)
        pyautogui.press("f1")
        self.console.print(f'✅ List load complete')
        time.sleep(0.5)

    def fill_rows_in_journal_screen(self, rows):
        text = ""
        for row in rows:
            row.pop(0)
            self.focus_on_screen()
            for value in row:
                text += str(value) + "\r\n"
            # blank fields
            for _ in range(3):
                text += "" + "\r\n"
            # c.copy(text)
        pc.copy(text)
        pc.waitForPaste()
        time.sleep(1)
        self.focus_on_screen(False)
        pyautogui.click(button='right')
        time.sleep(0.025)
        pyautogui.press('enter', presses=1, interval=0.25)

    def fill_row_in_journal_screen(self, row):
        row.pop(0)
        self.focus_on_screen()
        text = ""
        for value in row:
            text += str(value) + "\r\n"
        # blank fields
        for _ in range(3):
            text += "" + "\r\n"
        # c.copy(text)
        pc.copy(text)
        pc.waitForPaste()
        time.sleep(1)
        self.focus_on_screen(False)
        pyautogui.click(button='right')
        time.sleep(0.025)
        pyautogui.press('enter', presses=1, interval=0.25)
        # pyautogui.press('right', presses=4, interval=0.25)
