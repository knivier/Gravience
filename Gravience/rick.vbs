Dim WshShell, objIE

Set WshShell = CreateObject("WScript.Shell")

WshShell.Run "https://www.youtube.com/watch?v=E4WlUXrJgy4", 1, False

' Wait for 10 seconds (10000 milliseconds)
WScript.Sleep 10000

' Try to close the default browser (works for many browsers, like Chrome, Edge, etc.)
WshShell.Run "taskkill /F /IM chrome.exe", 0, False ' For Chrome
WshShell.Run "taskkill /F /IM msedge.exe", 0, False ' For Edge
WshShell.Run "taskkill /F /IM firefox.exe", 0, False ' For Firefox

' Cleanup
Set WshShell = Nothing
