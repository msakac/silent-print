import os
import winreg
import tkinter as tk
from tkinter import messagebox
import winshell
from win32com.client import Dispatch


def find_chrome_path():
    possible_paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    ]

    for path in possible_paths:
        if os.path.exists(path):
            return path

    try:
        reg_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as key:
            chrome_path, _ = winreg.QueryValueEx(key, "")
            if os.path.exists(chrome_path):
                return chrome_path
    except FileNotFoundError:
        pass

    return None


def create_shortcut(chrome_path: str, shortcut_name: str):
    desktop = winshell.desktop()
    path = os.path.join(desktop, f"{shortcut_name}.lnk")
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = chrome_path
    shortcut.Arguments = "--kiosk-printing"
    shortcut.WorkingDirectory = os.path.dirname(chrome_path)
    shortcut.IconLocation = chrome_path
    shortcut.save()


def on_create_shortcut():
    chrome_path = find_chrome_path()
    if chrome_path:
        shortcut_name = entry.get()
        if shortcut_name:
            create_shortcut(chrome_path, shortcut_name)
            messagebox.showinfo(
                "Success", f"Shortcut '{shortcut_name}' created successfully.")
        else:
            messagebox.showwarning(
                "Input Error", "Please enter a name for the shortcut.")
    else:
        messagebox.showerror("Error", "Google Chrome is not installed.")


# Create the main window
root = tk.Tk()
root.title("Chrome Shortcut Creator")

# Create and place the label and entry for the shortcut name
label = tk.Label(root, text="Enter the name for the Chrome shortcut:")
label.pack(pady=10)
entry = tk.Entry(root, width=40)
entry.pack(pady=5)

# Create and place the button to create the shortcut
button = tk.Button(root, text="Create Shortcut", command=on_create_shortcut)
button.pack(pady=20)

# Run the GUI event loop
root.mainloop()
