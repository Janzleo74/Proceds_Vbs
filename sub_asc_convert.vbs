
Option Explicit

Dim objshell
Set objshell=CreateObject("wscript.shell")

Dim nome, miarray, letra, toasc, i, tam, c, timer
timer= 200

letra="d"

toasc=Asc(letra)
wscript.echo toasc

nome="Janz"
tam=Len(nome)
wscript.echo tam


miarray=Array(0)
Redim miarray(tam)

'OBS: Como as arrays sao 0_based, um qualquer valor que damos
'no Redim, sempre nos trará um "size+1" [-->Asc_Convert..]


For i=1 To tam
miarray(i)=Asc(Mid(nome,i,1))
Next

REM For c=0 To UBound(miarray)
REM MsgBox miarray(c)
REM Next

objshell.run "notepad.exe"
wscript.sleep 2000

For c=1 To UBound(miarray)
wscript.sleep 250
objshell.sendkeys miarray(c)
wscript.sleep timer
objshell.sendkeys "-"
wscript.sleep timer
Next

REM wscript.sleep 6000
REM objshell.sendkeys "%asn"


