
/*
Author: João pedro Aguiar
Descrição:
Função que executa uma BAT que realiza a inicialização do programa Supervisor_Atendimento. 
A BAT precisa estar salva dentro da pasta TOTVS na unidade C do computador. 
Essa BAT é util para que o usuário possa executar aplicações com permissão de Admin.
*/
User Function Bat_Integração()

cPasta   := "C:\TOTVS\"
cArquivo := "RunasF.bat"

If MsgYesNo("Deseja iniciar o programa Supervisor_Atendimento?","Execução de BAT inicializador Supervisor.exe") 
	If  ShellExecute("OPEN", cArquivo, "", cPasta, 1) > 32
		MsgInfo ("Programa iniciado!", "Execução de BAT bem sucedida")
	Else 
		MsgAlert("Erro ao iniciar o programa Supervisor_Atendimento, entre em contato com a TI.", "Erro")
	EndIf
   
EndIf
	
Return
