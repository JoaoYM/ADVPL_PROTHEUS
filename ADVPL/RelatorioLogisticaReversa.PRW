
/*
+-----------+------------+----------------+-------------------+-------+---------------+
| Programa  | ADVPL >> EXCEL   | Desenvolvedor  |   Joao Pedro | Data  |   03/11/2022 |
+-----------+------------+----------------+-------------------+-------+---------------+
| Descricao | Gera relatorio em Excel com base nos dados inseridos pelo usuário       |
|           | no Pergunte.                                                            |
+-----------+-------------------------------------------------------------------------+
| Modulos   | -------                                                                 |
+-----------+-------------------------------------------------------------------------+
| Processos | Logistica Reversa                                                       |
+-----------+-------------------------------------------------------------------------+
|                  Modificacoes desde a construcao inicial                            |
+----------+-------------+------------------------------------------------------------+
| DATA     | PROGRAMADOR | MOTIVO                                                     |
+----------+-------------+------------------------------------------------------------+
*/

#INCLUDE 'Protheus.ch'
#INCLUDE 'Parmtype.ch'
#Include "TOTVS.ch"
#include "topconn.ch"



User Function ADVPL_To_EXCEL()

//---------------------Inicializa��o Vari�veis/Excel----------------------
	Local lRet

	Private cPasta      		:= GetTempPath()
	Private cWorkSheet			:= "Relatorio Logistica Reversa"
	Private cWork_2		        := "Notas de Devoluçao"
	Private cTable				:= "Relatorio Logistica"
	Private cTable_2			:= "Relatorio Logistica Notas de Devoluçao"
	Private nAlign 				:= 2
	Private nFormat 			:= 1
	Private lTotal 				:= .F.
	Private oExcel 				:= FWMsExcel():New()
	Private cArquivo    		:= GetTempPath() + ("Logistica_Reversa_" + DToS(DATE())) + ".xls"
	Private nPar 				:= 0
	Private cAliasTemporario_			:= GetNextAlias()
	Private cAliasTemporario_2			:= GetNextAlias()
    Private cAliasTemporario_3			:= GetNextAlias()



	if !MsgYesNo('<h1>Gostaria de gerar um novo relatorio?</h1>', "Confirma?")
		RETURN .f.
	ENDIF


	// Pergunte
	lRet := Pergunte("Name_Pergunte",.T.)
	if !lRet
		return NIL
	endif

	nPar := MV_PAR07

	

        If nPar = 3 .or. nPar = 2 
		
			// Query que buscará as tabelas e campos para gerar o relatório (VENDA DIRETA)
			BEGINSQL ALIAS "cAliasTemporario_3"
				SELECT
				 IIF(TableName_Vendas.Fieldname_Filial='0101', 'FARMACIA', 'INDUSTRIA') AS c_Filial, TableName_Vendas.Fieldname_Cliente, TableName_Vendas.Fieldname_DtEntregPrevist, TableName_Vendas.Fieldname_DtRecebimento, TableName_Vendas.Fieldname_ValorFrete, TableName_Vendas.Fieldname_Observacao, TableName_Vendas.Fieldname_Motivo, TableName_Vendas.Fieldname_LojaCliente, TableName_Vendas.Fieldname_N_Nota, TableName_Vendas.Fieldname_NomeCliente,
				
				 ISNULL(CONVERT(VARCHAR(8000),CONVERT(BINARY(8000),TableName_Vendas.Fieldname_Observacao)),"") AS Fieldname_Observacao,
				 ISNULL(CONVERT(VARCHAR(8000),CONVERT(BINARY(8000),TableName_Vendas.Fieldname_Motivo)),"") AS Fieldname_Motivo
				
				FROM
					%table:TableName_Vendas% AS TableName_Vendas

				WHERE
						TableName_Vendas.%NotDel%

					AND (
						TableName_Vendas.Fieldname_DtEntregPrevist BETWEEN %exp:DToS(MV_PAR01)% AND %exp:DToS(MV_PAR02)%
					)
					AND TableName_Vendas.Fieldname_DtRecebimento BETWEEN %exp:DToS(MV_PAR03)% AND %exp:DToS(MV_PAR04)%
					AND (
						TableName_Vendas.Fieldname_ValorFrete BETWEEN %exp:(MV_PAR05)% AND %exp:(MV_PAR06)%
					) 
			ENDSQL

			RptStatus({|| GeraRelatDocLeg(nPar)},"Aguarde","Gerando relatorio...")

        ENDIF
	
        If nPar = 3 .or. nPar = 1	
			// Query que buscará as tabelas e campos para gerar o relatório (PEDIDO DE VENDA)
			BEGINSQL ALIAS "cAliasTemporario_2"
				SELECT
				 IIF(TableName_Pedidos.Fieldname_Filial='0101', 'FARMACIA', 'INDUSTRIA') AS c_Filial, TableName_Vendas.Fieldname_Cliente, TableName_Vendas.Fieldname_DtEntregPrevist, TableName_Vendas.Fieldname_DtRecebimento, TableName_Vendas.Fieldname_ValorFrete, TableName_Vendas.Fieldname_Observacao, TableName_Vendas.Fieldname_Motivo, TableName_Vendas.Fieldname_LojaCliente, TableName_Vendas.Fieldname_N_Nota, TableName_Vendas.Fieldname_NomeCliente,
				
				 ISNULL(CONVERT(VARCHAR(8000),CONVERT(BINARY(8000),TableName_Pedidos.Fieldname_Observacao)),"") AS Observacao,
				 ISNULL(CONVERT(VARCHAR(8000),CONVERT(BINARY(8000),TableName_Pedidos.Fieldname_Motivo)),"") AS Motivo

				FROM
					%table:TableName_Pedidos% AS TableName_Pedidos
					INNER JOIN %table:TableName_DadosCliente% AS TableName_DadosCliente ON (TableName_Pedidos.Fieldname_NomeCliente = TableName_DadosCliente.FieldName_DadosCliente_Codig AND TableName_Pedidos.Fieldname_LojaCliente = TableName_DadosCliente.FieldName_DadosCliente_LOJA AND TableName_DadosCliente.%NotDel%)
					
				WHERE
					TableName_Pedidos.%NotDel%
					
					AND (
						TableName_Pedidos.Fieldname_DtEntregPrevist BETWEEN %exp:DToS(MV_PAR01)% AND %exp:DToS(MV_PAR02)%
					)
					AND TableName_Pedidos.Fieldname_DtRecebimento BETWEEN %exp:DToS(MV_PAR03)% AND %exp:DToS(MV_PAR04)%
					AND (
						TableName_Pedidos.Fieldname_ValorFrete BETWEEN %exp:(MV_PAR05)% AND %exp:(MV_PAR06)%
					)
			ENDSQL

				IF nPar = 3
					RptStatus({|| GeraRelatorio_Log_Reversa(nPar)},"Aguarde","Gerando relatorio...")
				ELSE 
					RptStatus({|| GeraRelatDocLeg(nPar)},"Aguarde","Gerando relatorio...")
				ENDIF
		EndIf

RETURN

Static Function GeraRelatDocLeg(nPar)

	Local cRotina := ""

	IF nPar = 2
	cAliasTemporario_3->(dbGoTop())
	SetRegua(0)
	ENDIF

	//-----------------------Criação Tabela----------------------------
	oExcel:AddworkSheet(cWorkSheet)
	oExcel:AddTable(cWorkSheet, cTable)
	oExcel:AddColumn(cWorkSheet,cTable,"Rotina",nAlign,nFormat,lTotal)
	oExcel:AddColumn(cWorkSheet,cTable,"Filial",nAlign,nFormat,lTotal)
	oExcel:AddColumn(cWorkSheet,cTable,"N�_Pedido",nAlign,nFormat,lTotal)
	oExcel:AddColumn(cWorkSheet,cTable,"N�_Cliente",nAlign,nFormat,lTotal)
	oExcel:AddColumn(cWorkSheet,cTable,"Nome Cliente",nAlign,nFormat,lTotal)
	oExcel:AddColumn(cWorkSheet,cTable,"Loja Cliente",nAlign,nFormat,lTotal)
	oExcel:AddColumn(cWorkSheet,cTable,"Data de Retirada",nAlign,nFormat,lTotal)
	oExcel:AddColumn(cWorkSheet,cTable,"Data de Recebimento",nAlign,nFormat,lTotal)
	oExcel:AddColumn(cWorkSheet,cTable,"Valor do Frete",nAlign,nFormat,lTotal)
	oExcel:AddColumn(cWorkSheet,cTable,"Observa��o",nAlign,nFormat,lTotal)
	oExcel:AddColumn(cWorkSheet,cTable,"Motivo",nAlign,nFormat,lTotal)


		IF nPar = 1 // Redireciona a execução do código para realizar a consulta na TableName_Pedidos
			RptStatus({|| GeraRelatorio_Log_Reversa(nPar)},"Aguarde","Gerando relatório...")
		ELSE  // END IF

	    //---------------------Alimentao tabela-----------------

		// Percorre dados obtidos da consulta SQL
		WHILE !cAliasTemporario_3->(EOF())
			
			cRotina  := "Venda Direta"

			oExcel:AddRow(cWorkSheet,cTable,{;
			cRotina,cAliasTemporario_3->c_Filial, cAliasTemporario_3->Fieldname_N_Nota, cAliasTemporario_3->Fieldname_Cliente, cAliasTemporario_3->Fieldname_NomeCliente, cAliasTemporario_3->Fieldname_LojaCliente, DTOC(STOD(cAliasTemporario_3->Fieldname_DtEntregPrevist)), DTOC(STOD(cAliasTemporario_3->Fieldname_DtRecebimento)), TRANSFORM(cAliasTemporario_3->Fieldname_ValorFrete, "@E 9,999.99"),;
			ALLTRIM(cAliasTemporario_3->Fieldname_Observacao), ALLTRIM(cAliasTemporario_3->Fieldname_Motivo)})
			cAliasTemporario_3->(dbSkip())
		EndDo
		cAliasTemporario_3->(dbCloseArea())
		ENDIF

		IF nPar = 2

		if !MsgYesNo('<h1>Gerar relatorio Notas de Devolu��o?</h1>', "Confirma?")
		InicializacaoDoXLS( )	
		RETURN
	    ENDIF
		GoTo Line204
		ENDIF

return .T. // end Static Function

 Static Function GeraRelatorio_Log_Reversa(nPar)

 Local cRotina := ""

  cAliasTemporario_2->(dbGoTop())
	SetRegua(0)

	//---------------------Alimentação tabela-----------------
	WHILE !cAliasTemporario_2->(EOF()) 	// Percorre dados obtidos da consulta SQL

		 cRotina  := "Pedido de Venda"   

		 oExcel:AddRow(cWorkSheet,cTable,{;
		 cRotina,cAliasTemporario_2->c_Filial, cAliasTemporario_2->Fieldname_N_Nota,cAliasTemporario_2->Fieldname_Cliente, cAliasTemporario_2->FieldName_DadosCliente_NOME,cAliasTemporario_2->Fieldname_LojaCliente, DTOC(STOD(cAliasTemporario_2->Fieldname_DtEntregPrevist)), DTOC(STOD(cAliasTemporario_2->Fieldname_DtRecebimento)), TRANSFORM(cAliasTemporario_2->Fieldname_ValorFrete, "@E 9,999.99"),;
		 ALLTRIM(cAliasTemporario_2->Fieldname_Observacao), ALLTRIM(cAliasTemporario_2->Fieldname_Motivo)})
		 cAliasTemporario_2->(dbSkip())
   
   //				 IIF(TableName_Pedidos.Fieldname_Filial='0101', 'FARMACIA', 'INDUSTRIA') AS c_Filial, TableName_Vendas.Fieldname_Cliente, TableName_Vendas.Fieldname_DtEntregPrevist, TableName_Vendas.Fieldname_DtRecebimento, TableName_Vendas.Fieldname_ValorFrete, TableName_Vendas.Fieldname_Observacao, TableName_Vendas.Fieldname_Motivo, TableName_Vendas.Fieldname_LojaCliente, TableName_Vendas.Fieldname_N_Nota, TableName_Vendas.Fieldname_NomeCliente,

   
    EndDo

	cAliasTemporario_2->(dbCloseArea())

    if !MsgYesNo('<h1>Gerar relatorio Notas de Devoluçao?</h1>', "Confirma?")
			InicializacaoDoXLS( )		
		    RETURN
	ENDIF

	//-----------------------Criação Tabela----------------------------
	oExcel:AddworkSheet(cWork_2)
	oExcel:AddTable(cWork_2, cTable_2)
	oExcel:AddColumn(cWork_2,cTable_2,"Filial",nAlign,nFormat,lTotal)
	oExcel:AddColumn(cWork_2,cTable_2,"N� Doc",nAlign,nFormat,lTotal)
	oExcel:AddColumn(cWork_2,cTable_2,"N� Serie",nAlign,nFormat,lTotal)
	oExcel:AddColumn(cWork_2,cTable_2,"Data Emissao",nAlign,nFormat,lTotal)
	oExcel:AddColumn(cWork_2,cTable_2,"Cod. Fornecedor",nAlign,nFormat,lTotal)
	oExcel:AddColumn(cWork_2,cTable_2,"Loja Cliente",nAlign,nFormat,lTotal)
	oExcel:AddColumn(cWork_2,cTable_2,"Nome Cliente",nAlign,nFormat,lTotal)
	oExcel:AddColumn(cWork_2,cTable_2,"Valor Bruto",nAlign,nFormat,lTotal)
	oExcel:AddColumn(cWork_2,cTable_2,"Valor do Frete",nAlign,nFormat,lTotal)
	oExcel:AddColumn(cWork_2,cTable_2,"Valor (Bruto - Frete)",nAlign,nFormat,lTotal)


    //Alerta Preenchimeno Pergunte

    MsgAlert("Favor informar o intervalo de tempo para as notas desejadas", "Notas por Periodo")

    // Pergunte
	lRet := Pergunte("Name_PergunteV2",.T.)
	if !lRet
		return NIL
	endif

// Query que buscará as tabelas e campos para gerar o relatório (VENDA DIRETA)
  BEGINSQL ALIAS "cAliasTemporario_"
    SELECT
     IIF(TableName_Notas.FieldName_Notas_FILIAL='0101','FARMACIA','INDUSTRIA') AS FILIAL,TableName_Notas.FieldName_Notas_DOC, TableName_Notas.FieldName_Notas_SERIE, TableName_Notas.FieldName_Notas_EMISSAO, TableName_Notas.FieldName_Notas_FORNECE, TableName_Notas.FieldName_Notas_LOJA, TableName_DadosCliente.FieldName_DadosCliente_NOME, TableName_Notas.FieldName_Notas_VALBRUT, TableName_Notas.FieldName_Notas_FRETE, (TableName_Notas.FieldName_Notas_VALBRUT - TableName_Notas.FieldName_Notas_FRETE) AS VALLIQ
    FROM
      %table:TableName_Notas% AS TableName_Notas
      INNER JOIN %table:TableName_DadosCliente% AS TableName_DadosCliente ON (TableName_Notas.FieldName_Notas_FORNECE = TableName_DadosCliente.FieldName_DadosCliente_COD AND TableName_Notas.FieldName_Notas_LOJA = TableName_DadosCliente.FieldName_DadosCliente_LOJA AND TableName_DadosCliente.D_E_L_E_T_ <> '*')
    WHERE
      TableName_Notas.D_E_L_E_T_ <> '*' AND TableName_Notas.FieldName_Notas_TIPO = 'D'  AND TableName_Notas.FieldName_Notas_EMISSAO BETWEEN %exp:DToS(MV_PAR01)% AND %exp:DToS(MV_PAR02)%		
    ENDSQL

    cAliasTemporario_->(dbGoTop())
	SetRegua(0)

	//---------------------Alimentação tabela-----------------
   WHILE !cAliasTemporario_->(EOF()) 	// Percorre dados obtidos da consulta SQL

	oExcel:AddRow(cWork_2,cTable_2,{;
	cAliasTemporario_->FILIAL,cAliasTemporario_->FieldName_Notas_DOC, cAliasTemporario_->FieldName_Notas_SERIE,DToC(SToD(cAliasTemporario_->FieldName_Notas_EMISSAO)), cAliasTemporario_->FieldName_Notas_FORNECE,cAliasTemporario_->FieldName_Notas_LOJA,AllTrim(cAliasTemporario_->FieldName_DadosCliente_NOME), TRANSFORM(cAliasTemporario_->FieldName_Notas_VALBRUT,"@E 9,999.99"), TRANSFORM(cAliasTemporario_->FieldName_Notas_FRETE,"@E 9,999.99"),;
	TRANSFORM(cAliasTemporario_->VALLIQ, "@E 9,999.99")})
	cAliasTemporario_->(dbSkip())
   EndDo
   cAliasTemporario_->(dbCloseArea())
   InicializacaoDoXLS( )

RETURN  // end Static Function

Static Function InicializacaoDoXLS( )

    oExcel:Activate()
	oExcel:GetXMLFile(cArquivo)

// ******************** Abertura do arquivo gerado************************ 	
 ShellExecute("OPEN", cArquivo, "", cPasta, 1) //Abre arquivo previamente gerado
	
Return 
