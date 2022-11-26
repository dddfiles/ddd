Set WshShell = WScript.CreateObject("WScript.Shell")
Call WshShell.Run("%windir%\system32\notepad.exe")
do
WshShell.SendKeys "HaH"
Call WshShell.Run("%windir%\system32\notepad.exe")
loop