Option Explicit

Dim objShell, objReg, strKeyPath, strValueName
Set objShell = CreateObject("WScript.Shell")

' Registry key path and value name for desktop wallpaper
strKeyPath = "Control Panel\Desktop\"
strValueName = "Wallpaper"

' Set wallpaper to empty (none)
objShell.RegWrite "HKCU\" & strKeyPath & strValueName, "", "REG_SZ"

' Apply the changes immediately
objShell.Run "RUNDLL32.EXE user32.dll, UpdatePerUserSystemParameters", 1, True

' Cleanup
Set objShell = Nothing