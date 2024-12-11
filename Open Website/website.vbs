Set WshShell = CreateObject("WScript.Shell")

Set x = CreateObject("WScript.Shell")

x.SendKeys "^{ESC}" 
WScript.Sleep(500)

x.SendKeys "run"
WScript.Sleep(500)

x.SendKeys "{ENTER}"
WScript.Sleep(500)

x.SendKeys "https://www.youtube.com/watch?v=GXK5jRhO_dM"
WScript.Sleep(700)

x.SendKeys "{ENTER}"
WScript.Sleep(1000)

MsgBox "Completed Website Opened - USB Ninja Ducky"

' Coded by d0gu77
