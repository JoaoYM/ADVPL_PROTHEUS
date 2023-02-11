#INCLUDE 'Protheus.ch'
#INCLUDE 'Parmtype.ch'
#INCLUDE 'rwmake.ch'
#INCLUDE 'PRTOPDEF.ch'


USER FUNCTION Ac_Entregas(npar)
	Local nDataentp              := ""
	Local nDataentef             := ""
	Local cObsEntrega            := SPACE(263)
	LOCAL lRet		   		     := .T.

	nDataentp	  	:= TableName_Pedido->FieldName_Pedido
	nDataentef 		:= TableName_Pedido->FieldName_Pedido
	cObsEntrega		:= IIF(EMPTY(TableName_Pedido->FieldName_Pedido),SPACE(500),TableName_Pedido->FieldName_Pedido + SPACE(263))

	lRet := STRUCT_SCREEN(@nDataentp,@nDataentef,@cObsEntrega)
	
	If !lRet
		Return .F.
	Else

		dbSelectArea( "TableName_Pedido" )
		IF Reclock ("TableName_Pedido",.F.)				  		  
				replace TableName_Pedido->FieldName_Pedido with nDataentp
				replace TableName_Pedido->FieldName_Pedido with nDataentef
				replace TableName_Pedido->FieldName_Pedido with cObsEntrega		
   		 EndIf

	Endif
	
RETURN lRet
	
STATIC FUNCTION STRUCT_SCREEN(nDataentp,nDataentef,cObsEntrega)

	LOCAL oGet
	LOCAL oGet2
	Local oGet3
	LOCAL oDlg
	LOCAL oButton
	LOCAL lOK

	DEFINE MSDIALOG oDlg FROM 200,300 TO 500,720 PIXEL TITLE OemToAnsi("Acompanhamento LogÃ­stico")

		TSay():New(008,020,{|| OemToAnsi("Entrega Prevista: ")},oDlg,,,,,,.T.,CLR_BLUE,CLR_WHITE,050,008)                    
		oGet := TGet():New(009,073,{|u| IF(PCount() == 0,nDataentp,nDataentp:=u)},oDlg,072,010,,,CLR_BLACK,CLR_WHITE,,,,.T.,,,,,.F.,,,,"","nDataentp")

		TSay():New(024,020,{|| OemToAnsi("Entrega efetiva: ")},oDlg,,,,,,.T.,CLR_BLUE,CLR_WHITE,050,008)
		oGet2 := TGet():New(025,0074,{|u| IF(PCount() == 0,nDataentef,nDataentef:=u)},oDlg,070,010,,,CLR_BLACK,CLR_WHITE,,,,.T.,,,,,.F.,,,,"","nDataentef")                    

		Tsay():New(060,020,{|| OemToAnsi( "Observação: " )},oDlg,,,,,,.T.,CLR_BLUE,CLR_WHITE,050,008)
		oGet3 := TGet():New(050,075,{|u| IF(PCount() == 0,cObsEntrega,cObsEntrega:=u)},oDlg,0100,050,,,CLR_BLACK,CLR_WHITE,,,,.T.,,,,,.F.,,,,"","cObsEntrega")                    


		@ 110,040 BUTTON oButton PROMPT "Salvar" OF oDlg PIXEL ACTION (lOK := .T.,oDlg:End())
		@ 110,100 BUTTON oButton PROMPT "Cancelar" OF oDlg PIXEL ACTION (oDlg:End(),lOK := .F.)

	ACTIVATE MSDIALOG oDlg

RETURN lOK
