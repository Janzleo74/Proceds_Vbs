
Option Explicit

Dim objshell
Set objshell=CreateObject("wscript.shell")

Dim mistr, miarray, letra, i, tam, c, timer
timer= 250

REM letra="d"
REM toasc=Asc(letra)
'------
REM mistr="Janz"

MsgBox "Hi there! Digite seu nome ou uma palavra para conversao em AscDecimal: "
mistr=InputBox("Digite aqui; e sem espacos")

tam=Len(mistr)

MsgBox "Esse eh o tamanho do que voce digitou: " &"termo de " &tam &" bytes."

miarray=Array(0)
Redim miarray(tam)

'OBS: Como as arrays sao 0_based, um qualquer valor que damos
'no Redim, sempre nos trará um "size+1" [..Asc_Convert..]

For i=1 To tam
miarray(i)=Asc(Mid(mistr,i,1))
Next

REM For c=1 To UBound(miarray)
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

wscript.sleep 1500
MsgBox "Mto bem! Agora, faca bom uso dessa ferramenta. ^^",VbInformation
MsgBox "Grato pela utilizacao e referencia."

REM wscript.sleep 6000
REM objshell.sendkeys "%asn"

