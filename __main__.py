import time
import pandas as pd
import openpyxl
import unicodedata
import tkinter as tk
from tkinter import Listbox, messagebox
from pynput.keyboard import Controller, Key
import keyboard
import pyperclip
import configparser
import os
import shutil
import psutil
import threading
from typing import Optional

class MnemoChoiceApp:
    """
    A Tkinter-based application for interactive data entry with filtering capabilities
    and automatic typing support using keyboard emulation.
    """
    def __init__(self, config_file: str = "config.ini"):
        self.pynput_keyboard = Controller()
        self.window: Optional[tk.Tk] = None
        self.entry: Optional[tk.Entry]
        self.last_geometry: Optional[str] = None  # Save window geometry before hiding
        self.config_file = config_file
        self.config = configparser.ConfigParser()
        self.file_path: str = ""
        self.file_tab: str = ""
        self.file_col1: str = ""
        self.file_col2: str = ""
        self.shortcut: str = ""
        self.autokill: str = ""
        self.opacity: str = ""
        self.data: Optional[pd.DataFrame] = None

    def kill_existing_instance(self):
        """Kill any existing process instances of the application."""
        current_process = psutil.Process()
        target_name = "MnemoChoice"

        for process in psutil.process_iter(['pid', 'name', 'cmdline']):
            if process.info['pid'] != current_process.pid:
                process_name = process.info['name'] or ""
                if target_name in process_name:
                    try:
                        process.terminate()
                        process.wait(timeout=3)
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        pass

    def init_config(self):
        """Initialize variables from the configuration file."""
        self.config.read(self.config_file)
        try:
            self.file_path = self.config['FILE']['PATH']
            self.file_tab = self.config['FILE']['TAB']
            self.file_col1 = self.config['FILE']['COLUMN1']
            self.file_col2 = self.config['FILE']['COLUMN2']
            self.shortcut = self.config['KEYBOARD']['SHORTCUT']
            self.autokill = self.config['PROCESS']['AUTOKILL']
            self.opacity = self.config['UI']['OPACITY']
        except KeyError as e:
            raise ValueError(f"Missing configuration key: {e}")

    def check_and_copy_file(self) -> str:
        """
        Check if the source file exists and copy it locally if needed.
        Returns the path to the local copy.
        """
        local_filename = './mnemo.xlsx'
        timeout = 5

        def copy_file():
            if os.path.exists(self.file_path):
                if not os.path.exists(local_filename) or os.path.getmtime(self.file_path) > os.path.getmtime(local_filename):
                    shutil.copy2(self.file_path, local_filename)
            elif not os.path.exists(local_filename):
                messagebox.showerror("Error", "Data file is inaccessible.")
                raise FileNotFoundError("Data file not found.")

        thread = threading.Thread(target=copy_file)
        thread.start()
        thread.join(timeout)

        if thread.is_alive():
            messagebox.showerror("Error", "Operation timed out.")
            raise TimeoutError("File copy operation exceeded the time limit.")

        return local_filename

    def load_data(self, filepath: str) -> pd.DataFrame:
        """
        Load data from the specified Excel file and sheet.
        Filters the data to only include the relevant columns based on the configuration.
        """
        df = pd.read_excel(filepath, sheet_name=self.file_tab, header=None)

        header_row = None
        for i, row in df.iterrows():
            if self.file_col1 in row.values and self.file_col2 in row.values:
                header_row = i
                break

        if header_row is None:
            raise ValueError(f"Columns '{self.file_col1}' and '{self.file_col2}' not found in the file.")

        df = pd.read_excel(filepath, sheet_name=self.file_tab, header=header_row)

        if self.file_col1 not in df.columns or self.file_col2 not in df.columns:
            raise ValueError(f"Columns '{self.file_col1}' and '{self.file_col2}' not found in the data.")

        return df[[self.file_col1, self.file_col2]]

    @staticmethod
    def normalize_text(input_str: str) -> str:
        """
        Normalize text by removing accents and converting to lowercase.
        """
        nfkd_form = unicodedata.normalize('NFKD', input_str)
        return ''.join([c for c in nfkd_form if not unicodedata.combining(c)]).lower()

    def type_text(self, text: str):
        """Type text character by character, handling uppercase letters."""
        for char in text:
            if char.isupper():
                self.pynput_keyboard.press(Key.shift)
                self.pynput_keyboard.press(char.lower())
                self.pynput_keyboard.release(char.lower())
                self.pynput_keyboard.release(Key.shift)
            else:
                self.pynput_keyboard.press(char)
                self.pynput_keyboard.release(char)

    def open_window(self):
        """
        Create and display the Tkinter-based UI for user input and suggestions filtering.
        """
        def update_listbox():
            typed_word = self.normalize_text(self.entry.get())
            filtered_words = self.data[self.data[self.file_col1].apply(self.normalize_text).str.contains(typed_word)]

            listbox.delete(0, tk.END)
            for _, row in filtered_words.iterrows():
                listbox.insert(tk.END, f"{row[self.file_col1]} ({row[self.file_col2]})")
            if listbox.size() > 0:
                listbox.select_set(0)

        def on_select():
            selection = listbox.get(tk.ACTIVE)
            if selection:
                selected_equivalent = selection.split('(')[-1].strip(')')
                keyboard.press_and_release('alt+tab')
                if self.autokill == '1':
                    self.last_geometry = self.window.geometry()
                    self.window.withdraw()
                time.sleep(0.2)
                self.type_text(selected_equivalent)
                self.entry.delete(0, 'end')
                update_listbox()
        
        def on_copy(event=None):
            selection = listbox.get(tk.ACTIVE)
            if selection:
                selected_equivalent = selection.split('(')[-1].strip(')')
                pyperclip.copy(selected_equivalent)
                keyboard.press_and_release('alt+tab')

        def on_autokill_check():
            self.autokill = str(atk.get())
            self.config['PROCESS']['AUTOKILL'] = self.autokill
            with open('config.ini', 'w') as configfile:
                self.config.write(configfile)

        def on_focusIn(event):
            self.window.attributes('-alpha', 1.0)

        def on_focusOut(event):
            self.window.attributes('-alpha', float(self.opacity))

        self.window = tk.Tk()
        self.window.title("MnemoChoice")
        self.window.geometry("400x300")
        self.window.attributes('-topmost', True)
        self.window.bind('<FocusOut>', on_focusOut)
        self.window.bind('<FocusIn>', on_focusIn)
        self.window.bind('<Down>', lambda e: listbox.focus_set())
        self.window.bind('<Return>', lambda _: on_select())
        self.window.withdraw()

        top_frame = tk.Frame(self.window)
        top_frame.pack(fill=tk.X, pady=10)

        self.entry = tk.Entry(top_frame)
        self.entry.pack(side=tk.LEFT, padx=10, expand=True)

        atk = tk.IntVar(value=self.autokill)
        chk1 = tk.Checkbutton(top_frame, text='Auto quit', variable=atk, onvalue=1, offvalue=0, command=on_autokill_check)
        chk1.pack(side=tk.RIGHT, padx=10)

        listbox = Listbox(self.window)
        listbox.pack(fill=tk.BOTH, expand=True)
        listbox.bind('<Double-1>', lambda _: on_select())
        self.entry.bind('<KeyRelease>', lambda _: update_listbox())

        button = tk.Button(self.window, width=3, height=1, text="📄", command=on_copy)
        button.place(relx=0.9, rely=0.9, anchor=tk.CENTER)
    
        update_listbox()
        self.window.mainloop()

    def toggle_window(self):
        """Toggle visibility of the application window."""
        if self.last_geometry:
            self.window.geometry(self.last_geometry)
        self.window.deiconify()
        self.entry.focus_set()

    def run(self):
        """Main entry point for the application."""
        self.kill_existing_instance()
        self.init_config()
        keyboard.add_hotkey(self.shortcut, self.toggle_window)
        filepath = self.check_and_copy_file()
        self.data = self.load_data(filepath)
        self.open_window()


def main():
    app = MnemoChoiceApp()
    app.run()

if __name__ == "__main__":
    main()
