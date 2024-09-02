import os
import winreg
import winshell
from win32com.client import Dispatch


def find_chrome_path():
    possible_paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    ]

    for path in possible_paths:
        if os.path.exists(path):
            print("Found Chrome at", path)
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


def create_shortcut(chrome_path, shortcut_name):
    desktop = winshell.desktop()
    path = os.path.join(desktop, f"{shortcut_name}.lnk")
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = chrome_path
    shortcut.Arguments = "--kiosk-printing"
    shortcut.WorkingDirectory = os.path.dirname(chrome_path)
    shortcut.IconLocation = chrome_path
    shortcut.save()


def main():
    chrome_path = find_chrome_path()
    if chrome_path:
        shortcut_name = input("Enter the name for the Chrome shortcut: ")
        create_shortcut(chrome_path, shortcut_name)
        print("Shortcut created successfully")
    else:
        print("Chrome not found")


if __name__ == '__main__':
    main()
