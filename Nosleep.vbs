Dim i
i = Cdate(inputbox ("Time"))
Set wshShell =wscript.CreateObject("WScript.Shell")
do
wscript.sleep 5000
wshshell.sendkeys "{NUMLOCK}"
IF (time()>i) then
exit do
END IF
loop