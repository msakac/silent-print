import os
import winreg
import customtkinter as ctk
from tkinter import messagebox
import tkinter
import winshell
from win32com.client import Dispatch
from PIL import Image, ImageTk


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


def create_shortcut(browser_path: str, shortcut_name: str, url: str, icon_path: str):
    desktop = winshell.desktop()
    shortcut_path = os.path.join(desktop, f"{shortcut_name}.lnk")
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.Targetpath = browser_path
    shortcut.Arguments = f"--kiosk-printing {url}"
    shortcut.WorkingDirectory = os.path.dirname(browser_path)
    shortcut.IconLocation = icon_path
    shortcut.save()


def on_create_shortcut():
    selected_browser = browser_var.get()
    selected_icon = icon_var.get()
    if selected_browser in found_browsers:
        browser_path = found_browsers[selected_browser]
        shortcut_name = entry_shortcut_name.get()
        url = entry_url.get()
        if shortcut_name and url:
            icon_path = os.path.abspath(os.path.join("assets", selected_icon))
            create_shortcut(browser_path, shortcut_name, url, icon_path)
            messagebox.showinfo(
                "Uspjeh", f"Prečac '{shortcut_name}' je uspješno kreiran.")
        else:
            messagebox.showwarning(
                "Greška u unosu", "Molimo unesite ime za prečac i link.")
    else:
        messagebox.showerror(
            "Greška", "Odabrani preglednik nije instaliran na Vašem računalu.")


# Create the main window
window = ctk.CTk()
window.title("Kreiranje prečaca preglednika za ispis")
window.resizable(False, False)
WINDOW_HEIGHT = 750
WINDOW_WIDTH = 600

screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()

x_cordinate = int((screen_width/2) - (WINDOW_WIDTH/2))
y_cordinate = int((screen_height/2) - (WINDOW_HEIGHT/2))

window.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{x_cordinate}+{y_cordinate}")

# Find browser paths
found_browsers = find_browser_paths()

# Display found browsers and their paths
browser_list_label = ctk.CTkLabel(
    window,
    text="Instalirani preglednici na računalu: ",
    font=("Arial", 13, "bold"),
)

browser_list_label.pack()
for browser, path in found_browsers.items():
    ctk.CTkLabel(window, text=f"{browser}", font=(
        "Arial", 13, "bold"), text_color="green").pack(pady=0)
    ctk.CTkLabel(
        window,
        text=f"Putanja: {path}",
        font=("Arial", 11, "italic"),
        text_color="gray"
    ).pack()
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
label.pack(pady=(20, 0))

browser_var = ctk.StringVar(window)
initial_browser = next(iter(found_browsers.keys()), "Odaberite preglednik")
browser_var.set(initial_browser)
browser_dropdown = ctk.CTkOptionMenu(
    window, variable=browser_var, values=list(found_browsers.keys()))
browser_dropdown.pack()

# Create and place the label and entry_shortcut_name for the shortcut name
label = ctk.CTkLabel(
    window,
    text="2. Unesite ime prečaca: ",
    font=("Arial", 12, "bold"),
)
label.pack(pady=(20, 0))
entry_shortcut_name = ctk.CTkEntry(
    window, width=300, textvariable=ctk.StringVar(value="Kasa"))
entry_shortcut_name.pack()

# Create and place the label and entry_url for the URL
label = ctk.CTkLabel(
    window,
    text="3. Unesite link početne stranice: ",
    font=("Arial", 12, "bold"),
)
label.pack(pady=(20, 0))
entry_url = ctk.CTkEntry(window, width=300, textvariable=ctk.StringVar(
    value="https://www.google.com/"))
entry_url.pack()

# Create and place the dropdown for selecting the icon
label = ctk.CTkLabel(
    window,
    text="4. Odaberite ikonu: ",
    font=("Arial", 12, "bold"),
)
label.pack(pady=(20, 0))

icon_var = ctk.StringVar(window)
icon_var.set("pos_icon_1.ico")

# Load icons
icon_files = ["pos_icon_1.ico", "pos_icon_2.ico", "pos_icon_3.ico"]
icons = [ImageTk.PhotoImage(Image.open(os.path.abspath(
    os.path.join("assets", icon))).resize((30, 30))) for icon in icon_files]

# Create a frame to hold the icon options
icon_frame = ctk.CTkFrame(window)
icon_frame.pack()

# Create radio buttons for each icon
for i, icon_file in enumerate(icon_files):
    # Create a frame to hold the icon and radio button
    item_frame = ctk.CTkFrame(icon_frame)
    item_frame.pack(anchor="w", pady=5)

    # Add the icon to the frame
    ctk.CTkLabel(item_frame, image=icons[i], text="").pack(
        side="left", padx=(0, 10))

    # Add the radio button to the frame
    icon_radio = ctk.CTkRadioButton(
        item_frame, text="", variable=icon_var, value=icon_file)
    icon_radio.pack(side="left")

# Create and place the button to create the shortcut
button = ctk.CTkButton(
    window,
    text="Kreiraj prečac",
    command=on_create_shortcut,
    font=("Arial", 14, "bold")
)
button.pack(pady=20)

# Run the GUI event loop
window.mainloop()
