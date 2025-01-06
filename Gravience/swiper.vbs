' Define the INI file path
Dim iniPath
iniPath = "C:\Temp\SystemInfo.ini"

' Create FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' Ensure the directory exists
If Not fso.FolderExists("C:\Temp") Then
    fso.CreateFolder("C:\Temp")
End If

' Create or overwrite the INI file
Set iniFile = fso.CreateTextFile(iniPath, True)

' Gather username
Dim username
Set objNetwork = CreateObject("WScript.Network")
username = objNetwork.UserName

' Write username to the INI file
iniFile.WriteLine "[SystemInfo]"
iniFile.WriteLine "Username=" & username

' Gather IP address
Dim ipConfig, ipRegex, matches, ip
Set ipConfig = CreateObject("WScript.Shell").Exec("cmd /c ipconfig")
Set ipRegex = New RegExp
ipRegex.Pattern = "IPv4 Address.*:\s*([\d\.]+)"
ipRegex.Global = True

Do Until ipConfig.StdOut.AtEndOfStream
    line = ipConfig.StdOut.ReadLine()
    If ipRegex.Test(line) Then
        Set matches = ipRegex.Execute(line)
        ip = matches(0).SubMatches(0)
        Exit Do
    End If
Loop

' Write IP address to the INI file
If Len(ip) > 0 Then
    iniFile.WriteLine "IPAddress=" & ip
Else
    iniFile.WriteLine "IPAddress=Not Found"
End If

' Close the file
iniFile.Close

' Notify completion (optional, silent otherwise)