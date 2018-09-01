
'Novo trejno_Proced_Backup

'HEAD ====================

Option Explicit

Dim ObjShell, ObjFso

Set ObjShell=CreateObject("wscript.shell")
Set ObjFso=CreateObject("scripting.FileSystemObject")

Dim fonto, destin, eko

'Execution_Body ==========

eko=MsgBox("Ola. Voce deseja fazer um backup agora?",VbYesNo+VbInformation)

If eko = VbYes Then
fonto=(InputBox("Digite aqui o caminho da pasta_origem: "))
MsgBox "Agora, digite o caminho da pasta_destino: ",VbInformation
destin=(InputBox("Endereco da pasta_destino: "))
wscript.sleep 1000

ObjFso.CopyFolder fonto,destin
MsgBox "Procedimento realizado com sucesso!"
wscript.sleep 1000
wscript.quit

Elseif eko = VbNo Then
MsgBox "Ok. Operacao cancelada.",VbCritical
End if


