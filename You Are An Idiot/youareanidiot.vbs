Set WshShell = CreateObject("WScript.Shell")
Set x = CreateObject("WScript.Shell")

x.SendKeys "^{ESC}"
WScript.Sleep(1000)

x.SendKeys "run"
WScript.Sleep(800)

x.SendKeys "{ENTER}"
WScript.Sleep(500)

x.SendKeys "https://www.youtube.com/watch?v=48rz8udZBmQ"
WScript.Sleep(700)

x.SendKeys "{ENTER}"
WScript.Sleep(1000)

x.SendKeys "{ENTER}"
WScript.Sleep(3000)

x.SendKeys "{F}"
WScript.Sleep(500)


