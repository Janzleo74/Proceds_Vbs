
'Simple_Proced_of_Backup_Dir
'Author: Leonardo La Janz.
'Period: 2sem/2018.

'HEAD ====================

Option Explicit

Dim ObjShell, ObjFso

Set ObjShell=CreateObject("wscript.shell")
Set ObjFSO=CreateObject("scripting.FileSystemObject")

Dim fonto, destin, eko

'Execution_Body ==========

eko=MsgBox("Ola. Voce deseja fazer um backup_dir agora?",VbYesNo+VbInformation)

If eko = VbYes Then
Do While fonto="" OR IsNumeric(fonto)
fonto=(InputBox("Digite o caminho da pasta_origem: "))
loop

MsgBox "Agora, digite o caminho da pasta_destino: ",VbInformation
Do While destin="" OR IsNumeric(fonto)
destin=(InputBox("Endereco da pasta_destino: "))
loop
wscript.sleep 1000

ObjFSO.CopyFolder fonto,destin
MsgBox "Procedimento realizado com sucesso!"
wscript.sleep 1000
wscript.quit

Elseif eko = VbNo Then
MsgBox "Ok. Operacao cancelada.",VbCritical
End if

