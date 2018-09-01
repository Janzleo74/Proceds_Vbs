
'ObjCom_Vars_Amb
'Simple program to extract the main Enviroment_Variables of the System.


'HEAD =========================================

Dim Objshell
Dim ObjFSO
Dim Op

Set Objshell=CreateObject("wscript.shell")
Set ObjFSO=CreateObject("scripting.FileSystemObject")


Op=MsgBox("Ola. Voce deseja um resumo do seu Sistema?",VbYesNo+VbQuestion)
if Op=VbYes Then
MsgBox "Ok. Vamos criar uma pasta para os arquivos(...)",VbOkOnly
wscript.sleep 2000

If ObjFSO.FolderExists("c:\resumo_sys01\") Then
MsgBox "Essa pasta com analise recente ja existe.",VbExclamation
wscript.sleep 2000
wscript.quit
Elseif Not ObjFSO.FolderExists("c:\resumo_sys01\") Then
ObjFSO.CreateFolder("c:\resumo_sys01\")
End if

wscript.sleep 2000
MsgBox "Certo. Pasta de referncia ja criada.",VbOkOnly
wscript.sleep 2000
MsgBox "Iniciando verificacao basica do OS(...)"
wscript.sleep 3500

Objshell.run "cmd.exe"
wscript.sleep 1500
Objshell.sendkeys "{enter}"
wscript.sleep 1500
Objshell.sendkeys "cd\"
wscript.sleep 1500
Objshell.sendkeys "{enter}"
wscript.sleep 1500
Objshell.sendkeys "c:"
wscript.sleep 1500
Objshell.sendkeys "{enter}"
wscript.sleep 1500
Objshell.sendkeys "cd c:\resumo_sys01\"
wscript.sleep 1500
Objshell.sendkeys "{enter}"
wscript.sleep 1500
Objshell.sendkeys "color 0a"
wscript.sleep 2000
Objshell.sendkeys "{enter}"
wscript.sleep 1500
Objshell.sendkeys "mode 90, 45"
wscript.sleep 1500
Objshell.sendkeys "{enter}"
wscript.sleep 2000

Objshell.sendkeys "echo Hoje eh dia: {%}date{%} > machina-infos1.txt"
wscript.sleep 1500
Objshell.sendkeys "{enter}"
wscript.sleep 1500
Objshell.sendkeys "echo Hora_Atual: {%}time{%} >> machina-infos1.txt"
wscript.sleep 1500
Objshell.sendkeys "{enter}"
wscript.sleep 2000
Objshell.sendkeys "echo Nome_da_Maquina: {%}computername{%} >> machina-infos1.txt"
wscript.sleep 1500
Objshell.sendkeys "{enter}"
wscript.sleep 1500
Objshell.sendkeys "echo Client: {%}homepath{%} >> machina-infos1.txt"
wscript.sleep 1500
Objshell.sendkeys "{enter}"
wscript.sleep 1500
Objshell.sendkeys "echo Processador: {%}processor_architecture{%} >> machina-infos1.txt"
wscript.sleep 1500
Objshell.sendkeys "{enter}"
wscript.sleep 1500
Objshell.sendkeys "echo :{%}processor_identifier{%} >> machina-infos1.txt"
wscript.sleep 1500
Objshell.sendkeys "{enter}"
wscript.sleep 1500
Objshell.sendkeys "echo {%}ComSpec{%} >> machina-infos1.txt"
wscript.sleep 1500
Objshell.sendkeys "{enter}"
wscript.sleep 1500
Objshell.sendkeys "echo."
wscript.sleep 1600
Objshell.sendkeys "{enter}"
wscript.sleep 1500
Objshell.sendkeys "echo:: Gerando arq2 SINTETIZADO(...)"
wscript.sleep 1600
Objshell.sendkeys "{enter}"
wscript.sleep 1600
Objshell.sendkeys "systeminfo > machina-infos2.txt"
wscript.sleep 1600
Objshell.sendkeys "{enter}"
wscript.sleep 7500
Objshell.sendkeys "exit"
wscript.sleep 1500
Objshell.sendkeys "{enter}"
wscript.sleep 1400

MsgBox "Abrindo a Pasta_de_Referencia(...)"
wscript.sleep 2000
Objshell.run "c:\resumo_sys01\"
wscript.sleep 2000
MsgBox "Agora, basta abrir os arquivos e conferir o resumo.",VbInformation
MsgBox "Grato pela confianca.",VbOkOnly

Elseif Op=VbNo Then
MsgBox "Tudo bem. Operacao encerrada.",VbCritical
End if

'-----

