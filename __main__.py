import pandas as pd
import openpyxl
import tkinter as tk
from tkinter import Listbox, messagebox
import keyboard
import pyperclip
import configparser
import os
import shutil
import psutil
import sys

# Fonction pour vérifier et tuer l'instance existante
def kill_existing_instance():
    current_process = psutil.Process()
    target_name = "MnemoChoice"  # Substring to search in process names

    for process in psutil.process_iter(['pid', 'name', 'cmdline']):
        if process.info['pid'] != current_process.pid:
            process_name = process.info['name'] or ""
            # Check if the target name appears in the process name
            if target_name in process_name:
                try:
                    process.terminate()
                    process.wait(timeout=3)  # Wait for termination, with a 3-second timeout
                except psutil.NoSuchProcess:
                    pass  # Ignore if the process has already exited
                except psutil.AccessDenied:
                    print(f"Access denied when trying to terminate process {process.info['pid']}")


# Initialisation des variables depuis le fichier de configuration
def init():
    global filePath, fileTab, fileCol1, fileCol2, shortcut
    config = configparser.ConfigParser()
    config.read('config.ini')

    filePath = config['FILE']['PATH']
    fileTab = config['FILE']['TAB']
    fileCol1 = config['FILE']['COLUMN1']
    fileCol2 = config['FILE']['COLUMN2']
    shortcut = config['KEYBOARD']['SHORTCUT']

# Vérification et copie du fichier si nécessaire
def check_and_copy_file():
    local_filename = './mnemo.xlsx'
    
    if os.path.exists(filePath):
        if not os.path.exists(local_filename) or os.path.getmtime(filePath) > os.path.getmtime(local_filename):
            shutil.copy2(filePath, local_filename)
    elif not os.path.exists(local_filename):
        messagebox.showerror("Erreur", "Fichier de données inaccessible.")
        raise FileNotFoundError("Fichier de données non trouvé.")
    
    return local_filename

# Charger les données depuis le fichier Excel avec les paramètres du fichier de config
def load_data(filepath):
    # Load the entire sheet without assuming header location
    df = pd.read_excel(filepath, sheet_name=fileTab, header=None)

    # Find the row with the specified headers
    header_row = None
    for i, row in df.iterrows():
        if fileCol1 in row.values and fileCol2 in row.values:
            header_row = i
            break

    # If header row not found, raise an error
    if header_row is None:
        raise ValueError(f"Columns '{fileCol1}' and '{fileCol2}' not found in the file.")

    # Load data starting from the row below the headers
    df = pd.read_excel(filepath, sheet_name=fileTab, header=header_row)

    # Filter to keep only relevant columns
    if fileCol1 not in df.columns or fileCol2 not in df.columns:
        raise ValueError(f"Columns '{fileCol1}' and '{fileCol2}' not found in data.")
    
    return df[[fileCol1, fileCol2]]


# Fonction pour afficher la fenêtre de saisie et filtrer les propositions
def open_window(data):
    def on_keypress(event):
        if event.keysym not in ["Up", "Down", "Return"]:
            entry.focus_set()  # Focus sur l'entrée si d'autres touches sont pressées
            typed_word = entry.get().lower()
            filtered_words = data[data[fileCol1].str.lower().str.contains(typed_word)]
            update_listbox(filtered_words)

    def update_listbox(filtered_words):
        listbox.delete(0, tk.END)
        for index, row in filtered_words.iterrows():
            listbox.insert(tk.END, f"{row[fileCol1]} ({row[fileCol2]})")
        if listbox.size() > 0:
            listbox.select_set(0)  # Sélectionne le premier élément
            
    def on_entry_return(event):
        # Check if there's exactly one item in the listbox when Enter is pressed in the entry
        if listbox.size() == 1:
            listbox.select_set(0)
            on_select()  # Automatically select the single item

    def on_select(event=None):
        selection = listbox.get(tk.ACTIVE)
        if selection:
            selected_equivalent = selection.split('(')[-1].strip(')')
            pyperclip.copy(selected_equivalent)
            window.destroy()
            keyboard.press_and_release('ctrl+v')

    window = tk.Tk()
    window.title("Recherche de mot")
    window.geometry("400x300")
    window.lift()
    window.attributes('-topmost', True)
    window.after(50, lambda: window.focus_force())

    # Champ de saisie avec autofocus
    entry = tk.Entry(window)
    entry.pack(pady=10)
    entry.bind('<KeyRelease>', on_keypress)
    entry.bind('<Down>', lambda e: listbox.focus_set())  # Si flèche bas, focus sur listbox
    entry.bind('<Return>', on_entry_return)
    window.after(100, lambda: entry.focus())

    # Liste de propositions
    listbox = Listbox(window)
    listbox.pack(fill=tk.BOTH, expand=True)
    listbox.bind('<Return>', on_select)

    window.mainloop()

# Détecter la combinaison de touches
def on_trigger():
    try:
        filepath = check_and_copy_file()
        data = load_data(filepath)
        open_window(data)
    except Exception as e:
        messagebox.showerror("Erreur lors de l'activation", e)

# Démarrer le script en tuant l'instance existante, puis lancer les étapes principales
if __name__ == "__main__":
    kill_existing_instance()
    init()
    keyboard.add_hotkey(shortcut, on_trigger)
    keyboard.wait()