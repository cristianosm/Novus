#INCLUDE "rwmake.ch"
#INCLUDE "Protheus.ch"
#INCLUDE "TopConn.ch"

/*
†††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ƒκκκκκκκκκκθκκκκκκκκκκκκκκκκθκκκκκκθκκκκκκκκκκκκκκκκκθκκκκκκθκκκκκκκκκκκΘ±±
±±ΌPrograma  Ό GL_PREVIVENDAS ΌAutor ΌGuilherme Strich Ό Data Ό10/06/2014 Ό±±
±±νκκκκκκκκκκζκκκκκκκκκκκκκκκκζκκκκκκζκκκκκκκκκκκκκκκκκζκκκκκκζκκκκκκκκκκκ ±±
±±ΌDesc.     Ό Importacao de PREVISAO DE VENDAS 					      Ό±±
±±νκκκκκκκκκκζκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκ ±±
±±ΌEquipe    Ό SAD Global                                                 Ό±±
±±ικκκκκκκκκκζκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκ*±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
*/
********************************************************************

User Function GL_PREVIVENDAS()

	//| Desativa pelo Cristiano Machado... Solicitado pelo Jaime... Data: 18/11/2015 ... Utilizar SG_IMPPREV.PRW |
	Alert("Este Fonte Foi DESATIVADO " + ProcName() )

Return()

/*
	Local oTelaIni
	Local oPanel
	Private cArquivo1    :=	""
	Private cArquivo2    :=	""
	Private cArquivo3    :=	""
	Private oFile1 		:= 	Nil
	Private oDest 		:= 	Nil
	Private oConfirm 	:= 	Nil
	Private cCaminho1	:= ""
	Private cDest       := "SB1"
	Private cAnoData 	:= "2015"
		//Alert("Opera‹o realizada com sucesso")
	IF !MsgYesNo("Deseja importar um arquivo ? ")
		FATA050()
		Return
	EndIf

	@ 200,001 TO 430,300 DIALOG oTelaIni TITLE 'Importa‹o de Dados de Previs‹o de Vendas'
	@ 002,002 TO 090,150 OF oTelaIni Pixel
	@ 005,008 Say "Arquivo de Importa‹o"  OF oTelaIni Pixel
	@ 014,008 MSGET  oFile1 VAR cArquivo1 WHEN .F. SIZE 100,09  OF oTelaIni Pixel
	@ 014,114 BUTTON "Abrir" SIZE 30,11  ACTION (fSelecFile(@cArquivo1,@cCaminho1)) OF oTelaIni Pixel
	@ 040,008 Say  "Ano: " OF oTelaIni Pixel
	@ 040,040 Get cAnoData SIZE 30,11 Picture "@!" OF oTelaIni Pixel
	@ 095,085 BUTTON "Cancelar" 	SIZE 30,12  ACTION (oTelaIni:End()) Message "Cancelar Importa‹o"  OF oTelaIni Pixel
	@ 095,118 BUTTON oConfirm PROMPT "Confirmar" 	SIZE 30,12 	ACTION ( IIF(fValida(),oTelaIni:End(),) ) Message "Confirmar Importa‹o"  OF oTelaIni Pixel
	Activate Dialog oTelaIni Centered
Return

/*†††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††
ƒκκκκκκκκκκ„κκκκκκκκκκθκκκκκκκ„κκκκκκκκκκκκκκκκκκκκθκκκκκκ„κκκκκκκκκκκκκΘ
ΌFuncao    *fSelecFileΌAutor  *Tiago Welter        Ό Data *  06/05/11   Ό
νκκκκκκκκκκ―κκκκκκκκκκζκκκκκκκμκκκκκκκκκκκκκκκκκκκκζκκκκκκμκκκκκκκκκκκκκ
ΌDesc.     *Funcao que seleciona o arquivo para importacao.             Ό
ικκκκκκκκκκμκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκ*
§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§//
Static Function fSelecFile(cArquivo,cCaminho)
	Local cTipoArq	:= "Arquivo (*.CSV)        | *.CSV | "
	cCaminho		:= cGetFile(cTipoArq,"Escolha o caminho do Arquivo")

	For nI:=Len(cCaminho) TO 1 Step -1
		If SubStr(cCaminho,nI,1) == "\"
			cArquivo :=	SubsTr(cCaminho,nI+1,Len(cCaminho))
			Exit
		EndIf
	Next nI
Return

/*†††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††
ƒκκκκκκκκκκ„κκκκκκκκκκθκκκκκκκ„κκκκκκκκκκκκκκκκκκκκθκκκκκκ„κκκκκκκκκκκκκΘ
ΌFuncao    *fValida   ΌAutor  *Tiago Welter        Ό Data *  06/05/11   Ό
νκκκκκκκκκκ―κκκκκκκκκκζκκκκκκκμκκκκκκκκκκκκκκκκκκκκζκκκκκκμκκκκκκκκκκκκκ
ΌDesc.     *Validacoes.                                                 Ό
ικκκκκκκκκκμκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκ*
§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§/
Static Function fValida()
	Local lRet := .T.
	If Empty(cArquivo1)
		If Empty(cArquivo1)
			Aviso("Escolha o arquivo de Importa‹o","Clique no bot‹o Abrir e selecione um arquivo.",{"OK"})
			Return .F.
		EndIf
	Else
		Processa({|| fExcImport()})
	EndIf
Return lRet

/*†††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††
ƒκκκκκκκκκκ„κκκκκκκκκκθκκκκκκκ„κκκκκκκκκκκκκκκκκκκκθκκκκκκ„κκκκκκκκκκκκκΘ
ΌPrograma  *fExcImportΌAutor  *Tiago Welter        Ό Data *  06/05/11   Ό
νκκκκκκκκκκ―κκκκκκκκκκζκκκκκκκμκκκκκκκκκκκκκκκκκκκκζκκκκκκμκκκκκκκκκκκκκ
Ό          *Funcao que executa a importacao.                            Ό
ικκκκκκκκκκμκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκκ*
§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§/
Static Function fExcImport
	Local nRegua	:= 0
	Local _nHa		:= Nil
	Local aContas	:= {}
	Local aCCusto	:= {}
	Local cLinha	:= ""
	Local cCgc 	    := ""
	Local cNumDoc	:= ""
	nColMes 		:= 4
	cVend 			:= ""
	cDoc 			:= ""

	//-ABRE O ARQUIVO TEXTO E SOBE PARA BUFFER
	If (_nHa := FT_FUse(AllTrim(Alltrim(cCaminho1))))== -1
		help(" ",1,"NOFILEIMPOR")
		Return
	EndIf

	nRegua := FT_FLASTREC()
	Procregua(nRegua)
	FT_FGOTOP()
	nPerReg  := 0
	nInc     := 100 / nRegua
	nCnt	 := 0

	While !FT_FEOF()
		nPerReg += nInc
		Incproc("Carregando Arquivo... "+Str(nPerReg,6,2)+"%")
		cLinha 	:= FT_FREADLN()
		nCnt++
		aTmp 	:= {}
		aTmp 	:= separa( cLinha,';',.T.) //Coluna 1: DE	Coluna 2: PARA

		If !( Len(aTmp)>0 )
			FT_FSKIP()
			Loop
		EndIf

		//Se a linha for a Linha 1 (cabealho) pula para a proxima
		If nCnt == 1
			FT_FSKIP()
			Loop
		EndIf

		//Aqui vou verificar qual o ultimo documento do ano das metas
		cQuery := " SELECT TOP 1 CT_DOC "
		cQuery += " FROM "+RetSqlName("SCT")
		cQuery += " WHERE D_E_L_E_T_ != '*' "
		cQuery += " AND  SUBSTRING(CT_DATA,1,4) = '"+cAnoData+"' "
 		TCQuery cQuery ALIAS "TMPSCT" NEW
		If TMPSCT->(!Eof())
        	cNumDoc := TMPSCT->CT_DOC
		Else
			//Se nao achei nenhum documento, busco o ultimo cadastrado e somo mais 1
			cQuery := " SELECT TOP 1 CT_DOC "
			cQuery += " FROM "+RetSqlName("SCT")
			cQuery += " WHERE D_E_L_E_T_ != '*' "
			cQuery += " AND  SUBSTRING(CT_DATA,1,4) = '"+cAnoData+"' "
			cQuery += " ORDER BY CT_DOC ASC "
	 		TCQuery cQuery ALIAS "TMPSCT2" NEW
			If TMPSCT2->(!Eof())
				cNumDoc := TMPSCT2->CT_DOC
				cNumDoc := Soma1(cNumDoc)
			EndIf
			TMPSCT2->(DbCloseArea())
		EndIf
		TMPSCT->(DbCloseArea())

   	    cQuery 	:= " SELECT * FROM "+RetSqlName("SA3")+" WHERE A3_COD = '"+Alltrim(aTmp[01])+"' "
   	    TCQuery cQuery ALIAS "TMPSCT3" NEW
   	    If TMPSCT3->(!EoF())
   	    	cVend :=TMPSCT3->A3_COD
   	    Endif
   	    TMPSCT3->(DbCloseArea())

   	    If Empty(cVend)
	   	    cQuery 	:= " SELECT * FROM "+RetSqlName("SA3")+" WHERE A3_NOME like '%"+Alltrim(aTmp[02])+"%' "
   	    	TCQuery cQuery ALIAS "TMPSCT3" NEW
   	    	If TMPSCT3->(!EoF())
   	    		cVend :=TMPSCT3->A3_COD
   	    	Endif
   	    	TMPSCT3->(DbCloseArea())
   	    EndIf

	 	cDoc := ""
 		cQuery := " SELECT TOP 1 CT_DOC FROM "+RetSqlName("SCT")+" WHERE D_E_L_E_T_ = '' "
  		cQuery += " ORDER BY CT_DOC DESC "
   		TCQuery cQuery ALIAS "TMPSCT2" NEW
   		If TMPSCT2->(!Eof())
   	   		cDoc := Soma1(TMPSCT2->CT_DOC)
   		EndIf
   		TMPSCT2->(DbCloseArea())

		If !Empty(cDoc) .AND. !Empty(cVend)
			nColMes := 3        // Ajustar quantas colunas tem antes dos valores.

			//VERIFICO MES A MES
			For i:= 1 To 12
				If i == 10
					cMes 		:= "10"
					cDataAtu 	:= cAnoData+"1001"
					cDataAt2 	:= "01"+"/10/"+cAnoData
				Else
					cMes 		:= Alltrim(Str(i))
					cDataAtu 	:= cAnoData+iif(Len(cMes)==1,"0"+cMes,cMes)+"01"
					cDataAt2 	:= "01"+"/"+iif(Len(cMes)==1,"0"+cMes,cMes)+"/"+cAnoData
				EndIf

				If Empty(aTmp[nColMes])
					nValMes 	:= 0
				Else
					nValMes 	:= StrTran(aTmp[nColMes],".","")
					nValMes 	:= StrTran(nValMes,",",".")
				EndIf

				//VERIFICA SE JA FOI LANCADO ESTE VALOR PARA O MESMO VENDEDOR E MESMO MζS
				cQuery := " SELECT * "
				cQuery += " FROM "+RetSqlName("SCT")
				cQuery += " WHERE D_E_L_E_T_ != '*' "
				cQuery += " AND CT_DATA = '"+cDataAtu+"' "
				cQuery += " AND CT_VALOR = '"+nValMes+"' "
				cQuery += " AND CT_VEND = '"+cVend+"' "
				TCQuery cQuery ALIAS "TMPSCT" NEW
				If TMPSCT->(Eof())
					cSeq := ""
					cQuery := " SELECT TOP 1 CT_SEQUEN "
					cQuery += " FROM "+RetSqlName("SCT")+" "
					cQuery += " WHERE CT_DOC = '"+cDoc+"' "
					cQuery += " AND D_E_L_E_T_ != '*' "
					cQuery += " ORDER BY  CT_SEQUEN DESC "
					TcQuery cQuery ALIAS "TMPSCT4" NEW
					If TMPSCT4->(Eof())
						cSeq := "001"
					Else
						cSeq := Soma1(TMPSCT4->CT_SEQUEN)
					EndIf
					TMPSCT4->(DbCloseArea())

					If RecLock("SCT",.T.)
						SCT->CT_DATA		:= CTOD(cDataAt2)
						SCT->CT_DESCRI 		:= "PREVISΜO DE VENDAS "+cVend+" "+cAnoData
						SCT->CT_DOC			:= cDoc
						SCT->CT_FILIAL 		:= xFilial("SCT")
						SCT->CT_MOEDA		:= 1
						SCT->CT_QUANT		:= 1
						SCT->CT_SEQUEN		:= cSeq
						SCT->CT_VALOR		:= val(nValMes)//vAl(nValMes)
						SCT->CT_VEND		:= cVend
						MsUnlock()
			  		EndIf
				EndIf
				TMPSCT->(DbCloseArea())
				nColMes++
			Next
		EndIf

		FT_FSKIP()
	EndDo

			//Alert("Opera‹o realizada com sucesso")
	IF MsgYesNo(" Opera‹o realizada com sucesso. Deseja importar outro arquivo ? ")
			u_GL_PREVIVENDAS()
		Else
			FATA050()
	EndIf

	FT_FUSE()
Return
*/