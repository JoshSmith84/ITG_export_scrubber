# Via AI assisted search when troubleshooting.
# Need to test/incorporate in main script
# For now, my sole "client" will simply make the reg change himself.


import winreg


def enable_long_paths():
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                             r"SYSTEM\CurrentControlSet\Control\FileSystem", 0,
                             winreg.KEY_ALL_ACCESS)
        winreg.SetValueEx(key, "LongPathsEnabled", 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(key)
        print(
            "Long path support enabled. You need to restart your computer for the changes to take effect.")
    except Exception as e:
        print("Error enabling long path support:", e)


if __name__ == "__main__":
    enable_long_paths()