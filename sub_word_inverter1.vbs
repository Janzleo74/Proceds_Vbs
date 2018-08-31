
'Inverter_v02

Option Explicit 

Dim objshell
Set objshell=CreateObject("wscript.shell")

Dim entry, tam, myarray, c, timer
Timer=200

MsgBox "Ola. Digite a palavra, nome ou frase que deseja inverter: "
entry=InputBox("Digite aqui, com um espaco para cada letra.", "","Assim-> c o m o  e s t e  e x e m p l o")

'entry=" L e o n a r d o  J a n z"

tam=Len(entry)
MsgBox "Esse eh o tamanho do que vc digitou: " &tam &" caracteres"

'*first, To array. ENTAO --> from Ubond to LBound.

myarray=Split(entry)

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

wscript.sleep 2500
MsgBox "Mto bem! Agora eh soh copiar e fechar o Notepad.",VbInformation
MsgBox "Agradecido pelo uso e por sua referencia."

REM wscript.sleep 5000
REM objshell.sendkeys "{enter 2}"
REM wscript.sleep 2000
REM objshell.sendkeys "%"
REM wscript.sleep 1500
REM objshell.sendkeys "asn"

