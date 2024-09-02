import os
import winreg
import customtkinter as ctk
from tkinter import messagebox
import tkinter
import winshell
from win32com.client import Dispatch

def find_browser_paths():
    browsers = {
        "Google Chrome": [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        ],
        "Microsoft Edge": [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
        ]
    }

    registry_paths = {
        "Google Chrome": r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
        "Microsoft Edge": r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe"
    }

    found_browsers = {}

    for browser, paths in browsers.items():
        for path in paths:
            if os.path.exists(path):
                found_browsers[browser] = path
                break
        else:
            # If no file path is found, check the registry
            try:
                reg_path = registry_paths[browser]
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as key:
                    browser_path, _ = winreg.QueryValueEx(key, "")
                    if os.path.exists(browser_path):
                        found_browsers[browser] = browser_path
            except FileNotFoundError:
                pass

    return found_browsers

def create_shortcut(browser_path: str, shortcut_name: str):
    desktop = winshell.desktop()
    path = os.path.join(desktop, f"{shortcut_name}.lnk")
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = browser_path
    shortcut.Arguments = "--kiosk-printing"
    shortcut.WorkingDirectory = os.path.dirname(browser_path)
    shortcut.IconLocation = browser_path
    shortcut.save()

def on_create_shortcut():
    selected_browser = browser_var.get()
    if selected_browser in found_browsers:
        browser_path = found_browsers[selected_browser]
        shortcut_name = entry.get()
        if shortcut_name:
            create_shortcut(browser_path, shortcut_name)
            messagebox.showinfo(
                "Uspjeh", f"Prečac '{shortcut_name}' za preglednik {selected_browser} je uspješno kreiran.")
        else:
            messagebox.showwarning(
                "Greška u unosu", "Molimo unesite ime za prečac.")
    else:
        messagebox.showerror("Greška", "Odabrani preglednik nije instaliran na Vašem računalu.")

# Create the main window
window = ctk.CTk()
window.title("Kreiranje prečaca preglednika za ispis")
window.resizable(False, False)
window_height = 400
window_width = 600

screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()

x_cordinate = int((screen_width/2) - (window_width/2))
y_cordinate = int((screen_height/2) - (window_height/2))

window.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))

# Find browser paths
found_browsers = find_browser_paths()

# Display found browsers and their paths
browser_list_label = ctk.CTkLabel(
    window,
    text="Pronađeni preglednici i njihove putanje:",
    font=("Arial", 14, "bold"),
)

browser_list_label.pack(pady=10)
for browser, path in found_browsers.items():
    browser_label = ctk.CTkLabel(window, text=f"{browser}: {path}")
    browser_label.pack()
if not found_browsers:
    label = ctk.CTkLabel(
        window,
        text="Nije pronađen nijedan preglednik na računalu.",
        text_color="red",
        font=("Arial", 12, "bold"),
    )
    label.pack()

# Create and place the dropdown for selecting the browser
label = ctk.CTkLabel(
    window,
    text="1. Odaberite preglednik: ",
    font=("Arial", 12, "bold"),
)
label.pack(pady=(30, 0))

browser_var = ctk.StringVar(window)
initial_browser = next(iter(found_browsers.keys()), "Odaberite preglednik")
browser_var.set(initial_browser)
browser_dropdown = ctk.CTkOptionMenu(window, variable=browser_var, values=list(found_browsers.keys()))
browser_dropdown.pack()

# Create and place the label and entry for the shortcut name
label = ctk.CTkLabel(
    window,
    text="2. Unesite ime prečaca: ",
    font=("Arial", 12, "bold"),
)
label.pack(pady=(30, 0))
entry = ctk.CTkEntry(window, width=300, textvariable=ctk.StringVar(value="Kasa"))
entry.pack()

# Create and place the button to create the shortcut
button = ctk.CTkButton(window, text="Kreiraj prečac", command=on_create_shortcut)
button.pack(pady=20)

# Run the GUI event loop
window.mainloop()
