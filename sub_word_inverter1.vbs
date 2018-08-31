'Inverter_v01

Option Explicit 

Dim objshell
Set objshell=CreateObject("wscript.shell")

Dim ynome, cutL, cutR, tam, myarray, c, timer
Timer=200

ynome=" L e o n a r d o "

REM cutL=Left(mynome,8)
REM cutR=Right(mynome,4)

tam=Len(ynome)
MsgBox "Your name_size: " &tam

'*first, To array. ENTAO --> from Ubond to LBound.

myarray=Split(ynome)

REM --> Word_Inverter:

Dim uband, lband, i

uband=UBound(myarray)
lband=LBound(myarray)

objshell.run "notepad.exe"
wscript.sleep 2400

objshell.sendkeys "{enter}"

for i=uband To lband Step -1
'MsgBox myarray(i)
objshell.sendkeys myarray(i)
wscript.sleep timer
objshell.sendkeys " "
wscript.sleep timer
Next
wscript.sleep 5000
objshell.sendkeys "{enter 2}"
wscript.sleep 2000
objshell.sendkeys "%"
wscript.sleep 1500
objshell.sendkeys "asn"



