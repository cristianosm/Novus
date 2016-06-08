#Include "Totvs.ch"

#Define NMaxC 6  //| Numero Maximo de Nivel existente no Plano de Contas...|

#Define CTAN		1	//| Posicao Array Conta Novus  |
#Define DCTN 	2 	//| Posicao Array Descricao Conta Novus |
#Define CTAO		3	//| Posicao Array Conta Origem  |
#Define DCTO 	4 	//| Posicao Array Descricao Conta Novus
#Define RECO		5	//| Posicao Array Recno Origem  |
#Define CTAD 	6 	//| Posicao Array Conta Destino |
#Define DCTD 	7 	//| Posicao Array Descricao Conta Novus
#Define RECD		8	//| Posicao Array Recno Destino |

#Define _ENTER CHR(10) + CHR(13)

/*****************************************************************************\
**---------------------------------------------------------------------------**
** FUNÇÃO   : NomeProg    | AUTOR : Cristiano Machado  | DATA : 02/06/2015   **
**---------------------------------------------------------------------------**
** DESCRIÇÃO: DE: PARA: do Plano Referencial Antigo(2014) para Novo(2015)    **
**---------------------------------------------------------------------------**
** USO      : Especifico para o cliente Novus                                **
**---------------------------------------------------------------------------**
**---------------------------------------------------------------------------**
**            ATUALIZACOES SOFRIDAS DESDE A CONSTRUCAO INICIAL.              **
**---------------------------------------------------------------------------**
**   PROGRAMADOR   |   DATA   |            MOTIVO DA ALTERACAO               **
**---------------------------------------------------------------------------**
**                 |          |                                              **
**                 |          |                                              **
\*---------------------------------------------------------------------------*/

*******************************************************************************
User Function DePara_Plano_Ref()
*******************************************************************************

	Private _cAliOr	:= "CVD" 							//| Plano de Contas Referenciais   |
	Private _cCpoCO	:= "CVD->CVD_CTAREF"  	//| Campo que contem a Conta Origem ...|

	Private _cCpoCD	:= "CVN->CVN_CTAREF"  	//| Campo que contem a Conta Destino ...|
	Private _cAliDe	:= "CVN"								//| Cadastro de Plando Referencial

	Private aContaQ	:= {}									//| Array que vai receber conta quebrada no delimitador .|
	Private nNConta	:= Len(aContaQ)  			//| Variavel armazena Numero de niveis da Conta |

	Private cContaN	:= ""									//| Contem o Numero da Conta Novus formato texto |
	Private cContaO	:= ""									//| Contem o Numero da Conta Origemem formato texto |
	Private cContaD	:= ""									//| Contem o Numero da Conta Destino em formato texto |

	Private nPRecOr	:= 0										//| Armazena a Posicao do CVD a ser Substituida |
	Private nPRecDs	:= 0										//| Armazena a Posicao do CVD a ser Substituida |


	/// ALTERAR TAMANHO CAMPO E CODIGO QUANDO FOR EXECUTAR EM UMA EMPRESA
	Private cCoPlaN	:= Pad("002", 3)  			//| Codigo do Plano Referencial NOVO |
	Private cCoPlaV	:= Pad("001", 3)  			//| Codigo do Plano Referencial VELHO |


	Private aDePara	:= {}									//| Armazena o De para ... Para gerar o Log |

	Private cCodSCt	:="01"									//| Codigo da SubConta... Ao qual sera utilizada para o Mergulho().|

	DbSelectaRea("CVN");DbSetOrder(2) //| Cadastro de Plando Referencial
	DbSelectaRea("CVD");DbSetOrder(2) //| Plano de Contas Referenciais

	Processa( {|| Inicia() }, "Aguarde...", "Analizando Conta "+ cContaO +" Destino "+ cContaD +"...",.F.)

	If Iw_MsgBox("Deseja efetuar a troca das contas ? ","Atencao", "YESNO")

		Iw_Msgbox("Por favor, efetuar backup da TABELA CVD !!!","Atencao.","ALERT")

		Processa( {|| AplicaDePara() }, "Aguarde...", "Alterando Contas...",.F.)

	EndIf

Return()
*******************************************************************************
Static Function Inicia() // inicia o Procesamento Geral ..
*******************************************************************************

				//| CVN_FILIAL + CVN_CODPLA + CVN_CTAREF |
	aAdd( aDePara , { "CONTA", "DESCRICAO", "CONTA_REF_O", "DESC_REF_O", "RECNO_REF_O", "CONTA_REF_D", "DESC_REF_D", "RECNO_REF_D"  } )

	DbSelectArea(_cAliOr);DbGotop()
	While !EOF()

		cContaN	:= CV(CVD->CVD_CONTA)

		nPRecOr := Recno()

		cContaO := SelConta(CV(&_cCpoCO.)) 	//| Converte a Conta Origem Texto em Array e Obtem o numero de Nivels|

		cContaD := Pesquisa()						//| Primeira Pesquisa para Localizar a conta Destino ....|

		SalvaLog() //

		IncProc("Analizando Conta "+ cContaO +" Destino "+ cContaD +"...")

		DbSelectarea(_cAliOr)
		DbSkip()

	EndDo

	ShowLog() //| Apresenta Log De Para...|

Return()
*******************************************************************************
Static Function SelConta(cConta) //| Converte a Conta Origem Texto em Array e Obtem o numero de Nivels|
*******************************************************************************

	aContaQ := StrTokArr ( cConta, "." )
	nNConta := Len(aContaQ)


Return(cConta)
*******************************************************************************
Static Function Pesquisa()//| Primeira Pesquisa para Localizar a conta Destino ....|
*******************************************************************************
	Local cContAux 	:= "" //| Armazena a conta Origem que vai ser Procurada no Novo Plano |
	Local cContAch 	:= "" //| Armazena a conta Origem que vai ser Encontrada no Novo Plano |
	Local cContNov		:= "" //| Armazena a conta Destino Posicionada no Novo Plano |

//DbSelectaRea(_cAliDe);DbGotop()
//| Segue ate o Menor Nivel Dentro do Novo Plano Apartir da Conta Antiga |
	For N := 1 To nNConta

		cContAux += aContaQ[N]
		DbSelectaRea(_cAliDe);DbGotop()

		If DbSeek(xFilial(_cAliDe) + cCoPlaN + cContAux,.F.)
			cContAch := cContAux
			cContNov := CV(&_cCpoCD.)
		Else
			Exit
		EndIf

		cContAux += "."

	Next

	cContNov := Mergulha(cContNov) //| Continua Descendo até o ultimo Nivel apartir de onde parou utilizando a Conta Origem.. |

Return(cContNov)
*******************************************************************************
Static Function Mergulha(cContNov) //| Verifica se ainda existe algum nivel mais Baixo para descer no Destino...|
*******************************************************************************
	Local cContAux := ""

	cContAux := 	SelConta(cContNov)

	For N := nNConta To NMaxC

		DbSelectaRea(_cAliDe);DbGotop()

		If DbSeek(xFilial(_cAliDe) + cCoPlaN + cContAux,.F.)
			cContNov := cContAux
			cContAux += "."+cCodSCt

			DbSelectaRea(_cAliDe)
			nPRecDs := Recno()

		Else
			Exit

		EndIf

	Next

Return(cContNov)
*******************************************************************************
Static Function SalvaLog()//| Salva LOG de Execucao|
*******************************************************************************
							 //POSICIONE( <cAlias>, <nOrdem>, <cChave>, <cCampo> )
	Local cDescConta			:= Posicione( "CT1" , 1 , xFilial("CT1") + cContaN 					, "CT1_DESC01" )
	Local cDescOrigem 		:= Posicione( "CVN" , 2 , xFilial("CVN") + cCoPlaV + cContaO 	, "CVN_DSCCTA" )
	Local cDescDestino 		:= Posicione( "CVN" , 2 , xFilial("CVN") + cCoPlaN + cContaD 	, "CVN_DSCCTA" )

	aadd( aDePara , { cContaN, CV(cDescConta), cContaO, CV(cDescOrigem), CV(nPRecOr), cContaD, CV(cDescDestino),CV(nPRecDs)  } )


Return()
*******************************************************************************
Static Function ShowLog()//| Monta o Log ....|
*******************************************************************************
	Local cTxt := ""

	For F := 1 To Len(adePara)

		For L := 1 To Len(adePara[F])

			cTxt +=  adePara[F][L] + ";"

		Next

		cTxt += _ENTER

	Next

	Mostra(cTxt)

Return()
*********************************************************************
Static Function Mostra( cTxt )
*********************************************************************
	__cFileLog := MemoWrite(Criatrab(,.F.)+".log",cTxt)

	Define FONT oFont NAME "Tahoma" Size 6,12
	Define MsDialog oDlgMemo Title "Leitura Concluida." From 3,0 to 340,550 Pixel

	@ 5,5 Get oMemo  Var cTxt MEMO Size 265,145 Of oDlgMemo Pixel
	oMemo:bRClicked := {||AllwaysTrue()}
	oMemo:oFont := oFont

	Define SButton  From 153,235 Type 1 Action oDlgMemo:End() Enable Of oDlgMemo Pixel

	Activate MsDialog oDlgMemo Center

Return()
*******************************************************************************
Static Function CV(xVar)
*******************************************************************************

Return(Alltrim(cValToChAr(xVar)))
*******************************************************************************
Static Function AplicaDePara()
*******************************************************************************

	Dbselectarea("CVD");DbSetOrder(2)
	Dbselectarea("CVN");DbSetOrder(2)

	For nA := 2 To Len(aDePara)

	  //	Alert( "RECNO CVD.."+aDePara[nA][RECO])
		DbSelectarea("CVD");DbGotop()
		DbGoto(Val(aDePara[nA][RECO]))

	//	Alert( "RECNO CVN.."+aDePara[nA][RECD])
		DbSelectarea("CVN");DbGotop()
		DbGoto(Val(aDePara[nA][RECD]))

   		DbSelectarea("CVD")
		RecLock("CVD",.F.)

		CVD->CVD_CODPLA := CVN->CVN_CODPLA //| Codigo do Plano Referencial |
		CVD->CVD_CTAREF := CVN->CVN_CTAREF //| Codigo da Conta Referencial |

		MsUnlock()

		IncProc()

	Next

Return()