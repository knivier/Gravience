' Create the WScript.Shell object to run the external script
Set WshShell = CreateObject("WScript.Shell")

' Get the current script's folder
strCurrentFolder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
' Build the path to taskii.vbs in the same folder
strTaskScript = strCurrentFolder & "\taskii.vbs"

' Create a FileSystemObject to check if the file exists
Set fso = CreateObject("Scripting.FileSystemObject")

' Check if taskii.vbs exists
If fso.FileExists(strTaskScript) Then
    ' If the file exists, run it
    On Error Resume Next  ' Ignore errors temporarily
    WshShell.Run "cscript //nologo """ & strTaskScript & """"
    
    ' Check if there was an error running the script
    If Err.Number <> 0 Then
        MsgBox "taskii.vbs target nullified", vbCritical, "Error"
    Else
        MsgBox "taskii.vbs taskmonned", vbInformation, "Task Complete"
    End If
    On Error GoTo 0 ' Reset error handling to default

Else
    ' If the file does not exist, display an error message
    MsgBox "taskii.vbs nullified", vbCritical, "Null"
End If

' Clean up
Set fso = Nothing
Set WshShell = Nothing