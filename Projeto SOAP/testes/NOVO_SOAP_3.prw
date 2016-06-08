#include "totvs.ch"
//------------------------------------------------------------------
//Exemplo de configuração de TGrid em array com navegação por linha
//------------------------------------------------------------------
#define GRID_MOVEUP       0
#define GRID_MOVEDOWN     1
#define GRID_MOVEHOME     2
#define GRID_MOVEEND      3
#define GRID_MOVEPAGEUP   4
#define GRID_MOVEPAGEDOWN 5
//------------------------------------------------------------------
//Valores para a propriedade nHScroll que define o comportamento da
//barra de rolagem horizontal
//------------------------------------------------------------------
#define GRID_HSCROLL_ASNEEDED   0
#define GRID_HSCROLL_ALWAYSOFF  1
#define GRID_HSCROLL_ALWAYSON   2

#Define RELVER 400
#Define RELHOR 400


// MeuGrid ( Classe para encapsular acesso ao componente TGrid )
//------------------------------------------------------------------------------
CLASS MeuGrid

	DATA oGrid
	DATA oFrame
	DATA oButtonsFrame
	DATA oButtonHome
	DATA oButtonPgUp
	DATA oButtonUp
	DATA oButtonDown
	DATA oButtonPgDown
	DATA oButtonEnd
	DATA aData
	DATA nLenData
	DATA nRecNo
	DATA nCursorPos
	DATA nVisibleRows
	DATA nFreeze
	DATA nHScroll

	METHOD New(oDlg) CONSTRUCTOR
	METHOD onMove( o,nMvType,nCurPos,nOffSet,nVisRows )
	METHOD isBof()
	METHOD isEof()
	METHOD ShowData( nFirstRec, nCount )
	METHOD ClearRows()
	METHOD AddColumn()
	METHOD DoUpdate()
	METHOD SelectRow(n)
	METHOD GoHome()
	METHOD GoEnd()
	METHOD GoPgUp()
	METHOD GoPgDown()
	METHOD GoUp(nOffSet)
	METHOD GoDown(nOffSet)
	METHOD SetCSS(cCSS)
	METHOD SetFreeze(nFreeze)
	METHOD SetHScrollState(nHScroll)
	METHOD Dialogo()
ENDCLASS

METHOD New(oDlg, aData) CLASS MeuGrid
	Local oFont

	::oFrame:= tPanel():New(0,0,,oDlg,,,,,,RELVER,RELHOR )
	::nRecNo:= 1
	::nCursorPos:= 0
	::nVisibleRows:= 14
    // Forçado para 1o ::GoEnd()
	::aData:= aData
	::nLenData:= Len(aData)
	::oGrid:= tGrid():New( ::oFrame )
	::oGrid:Align:= CONTROL_ALIGN_ALLCLIENT

    //oFont := TFont():New('Tahoma',,-32,.T.)
    //::oGrid:SetFont(oFont)
    //::oGrid:setRowHeight(50)

	::oButtonsFrame				:= tPanel():New(0,0,, ::oFrame,,,,,, RELVER,RELHOR,.F.,.T. )
	::oButtonsFrame:Align		:= CONTROL_ALIGN_RIGHT
	::oButtonHome					:= tBtnBmp():NewBar( "VCTOP.BMP",,,,, {||::GoHome()},,::oButtonsFrame )
	::oButtonHome:Align			:= CONTROL_ALIGN_TOP
	::oButtonPgUp					:= tBtnBmp():NewBar( "VCPGUP.BMP",,,,, {||::GoPgUp()},,::oButtonsFrame )
	::oButtonPgUp:Align			:= CONTROL_ALIGN_TOP
	::oButtonUp						:= tBtnBmp():NewBar( "VCUP.BMP",,,,,{||::GoUp(1)},,::oButtonsFrame )
	::oButtonUp:Align			:= CONTROL_ALIGN_TOP
	::oButtonEnd						:= tBtnBmp():NewBar( "VCBOTTOM.BMP",,,,, {||::GoEnd()},,::oButtonsFrame )
	::oButtonEnd:Align			:= CONTROL_ALIGN_BOTTOM
	::oButtonPgDown				:= tBtnBmp():NewBar( "VCPGDOWN.BMP",,,,, {||::GoPgDown()},,::oButtonsFrame )
	::oButtonPgDown:Align		:= CONTROL_ALIGN_BOTTOM
	::oButtonDown					:= tBtnBmp():NewBar( "VCDOWN.BMP",,,,, {||::GoDown(1)},,::oButtonsFrame )
	::oButtonDown:Align			:= CONTROL_ALIGN_BOTTOM

	::oGrid:addColumn( 1, "Código", 50, CONTROL_ALIGN_LEFT )
	::oGrid:addColumn( 2, "Descrição", 150, 0 )
	::oGrid:addColumn( 3, "Valor", 50, CONTROL_ALIGN_RIGHT )

	::oGrid:bCursorMove:= {|o,nMvType,nCurPos,nOffSet,nVisRows| ::onMove(o,nMvType,nCurPos,nOffSet,nVisRows) }
	::ShowData(1)
	::SelectRow( ::nCursorPos )
    // configura acionamento do duplo clique
	::oGrid:bLDblClick	:= {|| MsgStop("oi") }
	::oGrid:bRowLeftClick :=  {|| ::Dialogo() }


RETURN

METHOD isBof() CLASS MeuGrid
RETURN  ( ::nRecno==1 )
METHOD isEof() CLASS MeuGrid
RETURN ( ::nRecno==::nLenData )
METHOD GoHome() CLASS MeuGrid
	if ::isBof()
		return
	endif
	::nRecno = 1
	::oGrid:ClearRows()
	::ShowData( 1, ::nVisibleRows )
	::nCursorPos:= 0
	::SelectRow( ::nCursorPos )
RETURN
METHOD AddColumn() CLASS MeuGrid

	::oGrid:addColumn( 6, "Exemplo", 50, CONTROL_ALIGN_RIGHT )


RETURN
METHOD GoEnd() CLASS MeuGrid
	if ::isEof()
		return
	endif

	::nRecno:= ::nLenData
	::oGrid:ClearRows()
	::ShowData( ::nRecno - ::nVisibleRows + 1, ::nVisibleRows )
	::nCursorPos:= ::nVisibleRows-1
	::SelectRow( ::nCursorPos )
RETURN
METHOD GoPgUp() CLASS MeuGrid
	if ::isBof()
		return
	endif

    // força antes ir para a 1a linha da grid
	if ::nCursorPos != 0
		::nRecno -= ::nCursorPos
		if ::nRecno <= 0
			::nRecno:=1
		endif
		::nCursorPos:= 0
		::oGrid:setRowData( ::nCursorPos, {|o| { ::aData[::nRecno,1], ::aData[::nRecno,2], ::aData[::nRecno,3] } } )
	else
		::nRecno -= ::nVisibleRows
		if ::nRecno <= 0
			::nRecno:=1
		endif
		::oGrid:ClearRows()
		::ShowData( ::nRecno, ::nVisibleRows )
		::nCursorPos:= 0
	endif
	::SelectRow( ::nCursorPos )
RETURN
METHOD GoPgDown() CLASS MeuGrid
	Local nLastVisRow

	if ::isEof()
		return
	endif

    // força antes ir para a última linha da grid
	nLastVisRow:= ::nVisibleRows-1

	if ::nCursorPos!=nLastVisRow

		if ::nRecno+nLastVisRow > ::nLenData
			nLastVisRow:= ( ::nRecno+nLastVisRow ) - ::nLenData
			::nRecno:= ::nLenData
		else
			::nRecNo += nLastVisRow
		endif

		::nCursorPos:= nLastVisRow
		::oGrid:setRowData( ::nCursorPos, {|o| { ::aData[::nRecno,1], ::aData[::nRecno,2], ::aData[::nRecno,3] } } )
	else
		::oGrid:ClearRows()
		::nRecno += ::nVisibleRows

		if ::nRecno > ::nLenData
			::nVisibleRows = ::nRecno-::nLenData
			::nRecno:= ::nLenData
		endif

		::ShowData( ::nRecNo - ::nVisibleRows + 1, ::nVisibleRows )
		::nCursorPos:= ::nVisibleRows-1
	endif

	::SelectRow( ::nCursorPos )
RETURN

METHOD GoUp(nOffSet) CLASS MeuGrid
	Local lAdjustCursor:= .F.
	if ::isBof()
		RETURN
	endif
	if ::nCursorPos==0
		::oGrid:scrollLine(-1)
		lAdjustCursor:= .T.
	else
		::nCursorPos -= nOffSet
	endif
	::nRecno -= nOffSet

    // atualiza linha corrente
	::oGrid:setRowData( ::nCursorPos, {|o| { ::aData[::nRecno,1], ::aData[::nRecno,2], ::aData[::nRecno,3] } } )

	if lAdjustCursor
		::nCursorPos:= 0
	endif
	::SelectRow( ::nCursorPos )
RETURN
METHOD GoDown(nOffSet) CLASS MeuGrid
	Local lAdjustCursor:= .F.
	if ::isEof()
		RETURN
	endif

	if ::nCursorPos==::nVisibleRows-1
		::oGrid:scrollLine(1)
		lAdjustCursor:= .T.
	else
		::nCursorPos += nOffSet
	endif
	::nRecno += nOffSet

    // atualiza linha corrente
	::oGrid:setRowData( ::nCursorPos, {|o| { ::aData[::nRecno,1], ::aData[::nRecno,2], ::aData[::nRecno,3] } } )
	if lAdjustCursor
		::nCursorPos:= ::nVisibleRows-1
	endif
	::SelectRow( ::nCursorPos )
RETURN
METHOD onMove( oGrid,nMvType,nCurPos,nOffSet,nVisRows ) CLASS MeuGrid
	::nCursorPos:= nCurPos
	::nVisibleRows:= nVisRows

	if nMvType == GRID_MOVEUP
		::GoUp(nOffSet)
	elseif nMvType == GRID_MOVEDOWN
		::GoDown(nOffSet)
	elseif nMvType == GRID_MOVEHOME
		::GoHome()
	elseif nMvType == GRID_MOVEEND
		::GoEnd()
	elseif nMvType == GRID_MOVEPAGEUP
		::GoPgUp()
	elseif nMvType == GRID_MOVEPAGEDOWN
		::GoPgDown()
	endif
RETURN
METHOD ShowData( nFirstRec, nCount ) CLASS MeuGrid
	local i, nRec, ci
	DEFAULT nCount:=30

	for i=0 to nCount-1
		nRec:= nFirstRec+i
		if nRec > ::nLenData
			RETURN
		endif
		ci:= Str( nRec )
		cb:= "{|o| { Self:aData["+ci+",1], Self:aData["+ci+",2], Self:aData["+ci+",3] } }"
		::oGrid:setRowData( i, &cb )
	next i
RETURN
METHOD ClearRows() CLASS MeuGrid
	::oGrid:ClearRows()
	::nRecNo:=1
RETURN
METHOD DoUpdate() CLASS MeuGrid
	::nRecNo:=1
	::Showdata(1)
	::SelectRow(0)
RETURN
METHOD SelectRow(n) CLASS MeuGrid
	::oGrid:setSelectedRow(n)
RETURN
METHOD SetCSS(cCSS) CLASS MeuGrid
	::oGrid:setCSS(cCSS)
RETURN

METHOD SetFreeze(nFreeze) CLASS MeuGrid
	::nFreeze := nFreeze
	::oGrid:nFreeze := nFreeze
RETURN
METHOD SetHScrollState(nHScroll) CLASS MeuGrid
	::nHScroll 			:= nHScroll
	::oGrid:nHScroll 	:= nHScroll
RETURN
METHOD Dialogo( ) CLASS MeuGrid

	Local cGet1 := "Define variable value" // Variavel do tipo caracter
	Local nGet2 := 0 // Variável do tipo numérica
	Local dGet3 := Date() // Variável do tipo Data
	Local lHasButton := .T.

	DEFINE MSDIALOG oDlg TITLE "Picture test" FROM 000, 000  TO 100, 100 COLORS 0, 16777215 PIXEL

	oGet1 := TGet():New( 005, 009, { | u | If( PCount() == 0, cGet1, cGet1 := u ) },oDlg, ;
		060, 010, "!@",, 0, 16777215,,.F.,,.T.,,.F.,,.F.,.F.,,.F.,.F. ,,"cGet1",,,,lHasButton  )
	oGet2 := TGet():New( 020, 009, { | u | If( PCount() == 0, nGet2, nGet2 := u ) },oDlg, ;
		060, 010, "@E 999.99",, 0, 16777215,,.F.,,.T.,,.F.,,.F.,.F.,,.F.,.F. ,,"nGet2",,,,lHasButton  )
	oGet3 := TGet():New( 035, 009, { | u | If( PCount() == 0, dGet3, dGet3 := u ) },oDlg, ;
		060, 010, "@D",, 0, 16777215,,.F.,,.T.,,.F.,,.F.,.F.,,.F.,.F. ,,"dGet3",,,,lHasButton  )

	ACTIVATE MSDIALOG oDlg CENTERED


return




//------------------------------------------------------------------
User Function NOVO_SOAP_3()

	Local oDlg, aData:={}, i, oGridLocal, oEdit, nEdit:= 0
	Local oBtnAdd, oBtnClr, oBtnLoa

    // configura pintura da TGridLocal
	cCSS:= "QTableView{ alternate-background-color: #FFFFFF; background: #A9D0F5; selection-background-color: #2E9AFE }"

    // configura pintura do Header da TGrid
	cCSS+= "QHeaderView::section { background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #616161, stop: 0.5 #505050, stop: 0.6 #434343,  stop:1 #656565); color: white; padding-left: 4px; border: 1px solid #6c6c6c; }"

    // Dados
	for i:=1 to 100
		cCodProd:= StrZero(i,6)
		if i<3
            // inserindo imagem nas 2 primeiras linhas
			cProd:= "RPO_IMAGE=OK.BMP"
		else
			cProd:= 'Produto '+cCodProd
		endif

		cVal = Transform( 10.50, "@E 99999999.99" )
		AADD( aData, { cCodProd, cProd, cVal } )
	next

	DEFINE DIALOG oDlg FROM 0,0 TO (600), (600) PIXEL

	oGrid:= MeuGrid():New(oDlg,aData)
	oGrid:SetFreeze(2)
	oGrid:SetCSS(cCSS)
    //oGrid:SetHScrollState(GRID_HSCROLL_ALWAYSON) // Somente build superior a 131227A

    // Aplica configuração de pintura via CSSoGrid:SetCSS(cCSS)
	@ 210, 10 GET oEdit VAR nEdit OF oDlg PIXEL PICTURE "99999"

	@ 210, 070 BUTTON oBtnAdd PROMPT "Go" 			OF oDlg PIXEL ACTION oGrid:SelectRow(nEdit)
	@ 210, 100 BUTTON oBtnClr PROMPT "Clear" 	OF oDlg PIXEL ACTION oGrid:ClearRows()
	@ 210, 150 BUTTON oBtnLoa PROMPT "Update" 	OF oDlg PIXEL ACTION oGrid:DoUpdate()
	@ 210, 200 BUTTON oBtnLoa PROMPT "Columm" 	OF oDlg PIXEL ACTION oGrid:AddColumn()
	@ 260, 070 BUTTON oBtnLoa PROMPT "Dialogo"	OF oDlg PIXEL ACTION oGrid:DIALOGO()


	ACTIVATE DIALOG oDlg CENTERED

RETURN


