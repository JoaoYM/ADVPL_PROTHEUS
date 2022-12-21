
#INCLUDE 'Protheus.ch'
#INCLUDE 'Parmtype.ch'
#INCLUDE 'rwmake.ch'
#INCLUDE 'PRTOPDEF.ch'





USER FUNCTION Logistica_Reversa(nPar) // Atribuindo parametro "nPar" para lidar com redundancias das rotinas Pedidos e Vendas 


Local nDataRetirada     := ""
Local nDataRecebimento  := ""
Local nValorFrete       := ""
Local cObs_Logistica    := SPACE(250)
Local cMotivoLogRV      := SPACE(80)
Local cIndenizacao      := ""
LOCAL lRet		   		:= .T.
Local cNumOrc			:= ""
Local cDoc				:= ""
Local cSerie		    := ""


if nPar = 1  // nPar = 1 _ Pedidos
	
	If (__cUserId $ GETMV("Parametro_GrupoAllowed")) .and. Empty(TableName_Pedido->FieldName_Nota)
		MsgAlert("Pedido ainda não faturado. Não será possível informar acompanhamento de entrega.","Pedido não faturado")
		Return
	EndIf

	nDataRetirada	  	:= TableName_Pedido->FieldName_Ent.Prevista 
	nDataRecebimento	:= TableName_Pedido->FieldName_DataReceb
	nValorFrete 		:= TableName_Pedido->FieldName_ValorFrete
	cObs_Logistica		:= IIF(EMPTY(TableName_Pedido->FieldName_Observacao),SPACE(250),TableName_Pedido->FieldName_Observacao + SPACE(250))
	cMotivoLogRV		:= IIF(EMPTY(TableName_Pedido->FieldName_Motivo ),SPACE(250),TableName_Pedido->FieldName_Motivo  + SPACE(80))
    cNumOrc			    := TableName_Pedido->FieldName_N_nota
	cDoc			    := TableName_Pedido->FieldName_Nota
	cSerie			    := TableName_Pedido->FieldName_Serie
	cIndenizacao        := IIF (TableName_Pedido->FieldName_Indeniz_PG == "1", 'SIM', IIF (TableName_Pedido->FieldName_Indeniz_PG == "2", 'NAO', '') )


else		// nPar = 2 _  Vendas 


    If (__cUserId $ GETMV("Parametro_GrupoAllowed")) .and. Empty(TableName_Venda->FieldName_Doc)
		MsgAlert("Pedido ainda não faturado. Não será possível informar acompanhamento de entrega.","Pedido não faturado")
		Return
	EndIf

	nDataRetirada	  	:= TableName_Venda->FieldName_Ent.Prevista 
	nDataRecebimento	:= TableName_Venda->FieldName_DataReceb
	nValorFrete 		:= TableName_Venda->FieldName_ValorFrete
	cObs_Logistica		:= IIF(EMPTY(TableName_Venda->FieldName_Observacao),SPACE(250),TableName_Venda->FieldName_Observacao )
   	cMotivoLogRV		:= IIF(EMPTY(TableName_Venda->FieldName_Motivo),SPACE(250),TableName_Venda->FieldName_Motivo )
    cNumOrc			    := TableName_Venda->FieldName_N_nota
	cDoc			    := TableName_Venda->FieldName_Doc
	cSerie			    := TableName_Venda->FieldName_Serie
	cIndenizacao        := IIF (TableName_Venda->FieldName_Indeniz_PG == "1", 'SIM', IIF (TableName_Venda->FieldName_Indeniz_PG == "2", 'NAO', '') )



endif 

// Montando tela exibição -> usuário -> Campos "Data da Retirada ", "Data do Recebimento", "Valor do Frete", "Observação", "Motivo" 
lRet := STRUCT_SCREEN(@nDataRetirada,@nDataRecebimento,@nValorFrete,@cObs_Logistica,@cMotivoLogRV,@cNumOrc,@cDoc,@cSerie,@nPar,@cIndenizacao)

If !lRet  // SE erro ao estruturar tabela 
	Return .F.
Else     // Consulta ao banco de dados Pedidos
    If nPar = 1
        dbSelectArea( "TableName_Pedido" )
        IF Reclock ("TableName_Pedido",.F.)				  		  
            replace TableName_Pedido->FieldName_Ent.Prevista  with nDataRetirada
     		replace TableName_Pedido->FieldName_DataReceb  with nDataRecebimento
            replace TableName_Pedido->FieldName_ValorFrete with nValorFrete
            replace TableName_Pedido->FieldName_Observacao  with cObs_Logistica
            replace TableName_Pedido->FieldName_Motivo  with cMotivoLogRV	
			replace TableName_Pedido->FieldName_Indeniz_PG  with IIF (cIndenizacao == "SIM", '1', IIF (cIndenizacao == "NAO", '2', '') )
	
        Endif        
    
    Else // Consulta ao banco de dados Pedidos
       
        dbSelectArea( "TableName_Venda" )
        IF Reclock ("TableName_Venda",.F.)				  		  
            replace TableName_Venda->FieldName_Ent.Prevista   with nDataRetirada
			replace TableName_Venda->FieldName_DataReceb  with nDataRecebimento
            replace TableName_Venda->FieldName_ValorFrete  with nValorFrete
            replace TableName_Venda->FieldName_Observacao  with cObs_Logistica	
            replace TableName_Venda->FieldName_Motivo  with cMotivoLogRV	
			replace TableName_Venda->FieldName_Indeniz_PG  with IIF (cIndenizacao == "SIM", '1', IIF (cIndenizacao == "NAO", '2', '') )
        Endif
   		 
    Endif

EndIf
	
RETURN lRet
	
STATIC FUNCTION STRUCT_SCREEN(nDataRetirada,nDataRecebimento,nValorFrete,cObs_Logistica,cMotivoLogRV,_cNumOrc,_cDoc,_cSerie,_nPar,cIndenizacao)

	LOCAL oGet
	LOCAL oGet2
	LOCAL oGet3
	LOCAL oDlg
	LOCAL oButton
	LOCAL lOK   := .F.
	Local oMemo
    LOCAL cRps  := ""
	Local cNfse := ""
	LOCAL cDoc  := ""
	LOCAL aArea := GetArea()
	Local aItems:= {"","SIM","NAO"}

	
    cDoc := "Doc: " + _cDoc + " / Série: " + Alltrim(_cSerie)
	cCombo1:= aItems[1]

	If _nPar <> 1
		cRPs  := Posicione("TableName_Pedido",11,XFILIAL("TableName_Pedido") + _cNumOrc,"FieldName_Nota")
		cNfse := Posicione("TableName_NFSE",6,XFILIAL("TableName_NFSE") + cRps + 'TableName_NFSE_2',"FieldName_NFSE_2")
		cRps  := "RPS " + Alltrim(cRps) + " - NFSE: " + Alltrim(cNfse)
	ElseIf Alltrim(_cSerie) == 'TableName_NFSE_2'
		cNfse := Posicione("TableName_NFSE",6,XFILIAL("TableName_NFSE") + _cDoc + 'TableName_NFSE_2',"FieldName_NFSE_2")
		cDoc  += " - NFSE: " + Alltrim(cNfse) + BuscNfS(_cNumOrc)
		RestArea(aArea)
	EndIf

	DEFINE MSDIALOG oDlg FROM 200,300 TO 600,760 PIXEL TITLE OemToAnsi("Logistica reversa")
    // RPs e Doc ---------------------------------------------------------------------------------------------------------------------------------------------------------
        TSay():New(008,015,{|| OemToAnsi(cRPs + "   " +cDoc)},oDlg,,,,,,.T.,CLR_BLUE,CLR_WHITE,250,008) 
	// Data de retirada --------------------------------------------------------------------------------------------------------------------------------------------------
		TSay():New(024,015,{|| OemToAnsi("Data de retirada: ")},oDlg,,,,,,.T.,CLR_BLUE,CLR_WHITE,050,008)                    
		oGet := TGet():New(030,068,{|u| IF(PCount() == 0,nDataRetirada,nDataRetirada:=u)},oDlg,070,010,,,CLR_BLACK,CLR_WHITE,,,,.T.,,,,,.F.,,,,"","nDataRetirada")
	// Data do recebimento -----------------------------------------------------------------------------------------------------------------------------------------------
		TSay():New(047,015,{|| OemToAnsi("Data do recebimento: ")},oDlg,,,,,,.T.,CLR_BLUE,CLR_WHITE,055,010)                    
		oGet2 := TGet():New(049,068,{|u| IF(PCount() == 0,nDataRecebimento,nDataRecebimento:=u)},oDlg,070,010,,,CLR_BLACK,CLR_WHITE,,,,.T.,,,,,.F.,,,,"","nDataRecebimento")
	// Indenização -----------------------------------------------------------------------------------------------------------------------------------------------
	    TSay():New(022,150,{|| OemToAnsi("Indeniza��o Paga: ")},oDlg,,,,,,.T.,CLR_BLUE,CLR_WHITE,050,008)                    
		oCombo1 := TComboBox():New(029,150,{|u|if(PCount() == 0,cIndenizacao,cIndenizacao:=u)},;
        aItems,80,15,oDlg,,;
        ,,,,.T.,,,,,,,,,"cIndenizacao")
	// Valor do Frete ----------------------------------------------------------------------------------------------------------------------------------------------------
		TSay():New(070,015,{|| OemToAnsi("Valor do Frete: ")},oDlg,,,,,,.T.,CLR_BLUE,CLR_WHITE,050,008)
		oGet3 := TGet():New(070,068,{|u| IF(PCount() == 0,nValorFrete,nValorFrete:=u)},oDlg,070,010,"@E 9,999.99",,CLR_BLACK,CLR_WHITE,,,,.T.,,,,,.F.,,,,"","nValorFrete")                    
	// Observação --------------------------------------------------------------------------------------------------------------------------------------------------------
		Tsay():New(100,015,{|| OemToAnsi( "Observa��o: " )},oDlg,,,,,,.T.,CLR_BLUE,CLR_WHITE,050,008)
		@ 096, 068 GET oMemo VAR cObs_Logistica MEMO SIZE 150, 30 OF oDlg PIXEL
	// Motivo ------------------------------------------------------------------------------------------------------------------------------------------------------------
		Tsay():New(137,015,{|| OemToAnsi( "Informe o motivo: " )},oDlg,,,,,,.T.,CLR_BLUE,CLR_WHITE,050,008)
	    @ 135, 068 GET oMemo VAR cMotivoLogRV MEMO SIZE 150, 25 OF oDlg PIXEL

		//Especificações botão "Cancelar"
		@ 165,102 BUTTON oButton PROMPT "Cancelar" SIZE 28,15 OF oDlg PIXEL ACTION (oDlg:End(),lOK := .F.) 

					
			If (__cUserId $ GETMV("Parametro_GrupoAllowed"))
			  //Especificações botão "Salvar"
		      @ 165,068 BUTTON oButton PROMPT "Salvar" SIZE 28,15 OF oDlg PIXEL ACTION (lOK := .T.,oDlg:End())
			  @ 165,135 BUTTON oButton PROMPT "Emitir Relatorio" SIZE 85, 015 PIXEL OF oDlg ACTION ( U_ADVPL_To_EXCEL(Rotina_Param) )

			EndIf

	ACTIVATE MSDIALOG oDlg

RETURN lOK

Static Function BuscNfS(_cNumOrc)

	Local cAliTMP := GetNextAlias()
	Local cNfs    := ""

	BEGINSQL ALIAS cAliTMP
		SELECT 
			TOP 1 TableName_Pedido.FieldName_Nota
		FROM
			%Table:TableName_Pedido% TableName_Pedido
		WHERE
			TableName_Pedido.C5_ZNUMPED = %Exp:_cNumOrc% 
			AND TableName_Pedido.%notdel%
            AND TableName_Pedido.C5_FILIAL = %xFilial:TableName_Pedido%
	ENDSQL

	DbSelectArea(cAliTMP)
    (cAliTMP)->(DbGoTop())

	If !((cAliTMP)->(EOF()).AND.(cAliTMP)->(BOF()))
		cNfs := " / NFS: " + Alltrim((cAliTMP)->FieldName_Nota)
	EndIf

Return cNfs
