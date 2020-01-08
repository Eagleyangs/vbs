Set objShell = CreateObject("Wscript.Shell") 
do
WScript.Sleep 300000
objShell.SendKeys "{F5}"
WScript.Sleep 300000
objShell.SendKeys "{F5}"
loop 
