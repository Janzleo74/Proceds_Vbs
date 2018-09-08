
'Inverter_v02

Option Explicit 

Dim objshell, entry, resp
Set objshell=CreateObject("wscript.shell")

resp=MsgBox("Ola. Digite o nome ou a palavra/frase vc deseja inverter: ",VbOkCancel)
If resp=VbCancel then
MsgBox "Ok. Procedimento abortado."
Wscript.quit
Elseif resp=VbOk then
entry=InputBox("digite aqui, p. favor:")

VortInversa(entry)
End if

'-----------

Sub VortInversa(string)
Dim i, c, myarray, timer, tam, lband, rband
timer=200

tam=Len(string)

myarray=Array()
Redim myarray(tam)

For i=1 To tam
myarray(i)=Mid(string,i,1)
Next

lband=LBound(myarray)
rband=UBound(myarray)

objshell.run "notepad.exe"
wscript.sleep 2000
objshell.sendkeys "{enter}"
wscript.sleep 1200

For c=rband To lband Step -1
objshell.sendkeys myarray(c)
wscript.sleep timer
REM objshell.sendkeys " "
REM wscript.sleep timer
Next

wscript.sleep 2100
MsgBox "Mto bem! Agora eh soh copiar e fechar o Notepad.",VbInformation
MsgBox "Grato pelo uso e pela referencia."
End Sub

