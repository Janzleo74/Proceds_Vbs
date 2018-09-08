
'Simple_Proced: Asc_Converter
'Author.: Leonardo La Janz.
'Period: 2sem/2018.

Option Explicit

Dim objshell
Set objshell=CreateObject("wscript.shell")

Dim mistr 

MsgBox "Hi. Digite seu nome ou uma palavra para conversao em AscDecimal: "
Do While mistr="" OR IsNumeric(mistr)
mistr=InputBox("Digite aqui; e sem espacos")
Loop

Call AscKonvert(mistr)

'------------

Sub AscKonvert(string)
Dim miarray, i, tam, c, timer
timer=250

tam=Len(string)

miarray=Array(0)
Redim miarray(tam) 'tam of string

For i=1 To tam
miarray(i)=Asc(Mid(string,i,1))
Next

REM For c=1 To UBound(miarray)
REM MsgBox miarray(c)
REM Next

objshell.run "notepad.exe"
wscript.sleep 2000

objshell.sendkeys "{enter}"
For c=1 To UBound(miarray)
wscript.sleep timer
objshell.sendkeys miarray(c)
wscript.sleep timer
objshell.sendkeys "-"
wscript.sleep timer
Next

wscript.sleep 2000
MsgBox "Mto bem! Agora, faca bom uso dessa ferramenta. ^^",VbInformation
MsgBox "Grato pela utilizacao e referencia."

REM wscript.sleep 5000
REM objshell.sendkeys "%asn"

End sub
