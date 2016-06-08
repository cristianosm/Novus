
#Include "Totvs.ch"
#Include "RwMake.ch"

#Define 	TRUE    .T.
#Define 	FALSE   .F.

//| Posicao do Produto, Local e Primeira Quantidade no Array Cabec....
#Define _PRODUTO 1
#Define _LOCAL 	2
#Define _PRIQTD 	3

#Define _ENTER Chr(10) + Chr(13)

/*****************************************************************************\
**---------------------------------------------------------------------------**
** FUNCAO   : NomeProg    | AUTOR : JAIME  | DATA : 25/07/2005               **
**---------------------------------------------------------------------------**
** DESCRICAO: IMPORTA ARQUIVO CONVENIO PARA SB1 E SG1                        **
**---------------------------------------------------------------------------**
** USO      : Especifico para o cliente NOVUS                                **
**---------------------------------------------------------------------------**
**---------------------------------------------------------------------------**
**            ATUALIZACOES SOFRIDAS DESDE A CONSTRUCAO INICIAL.              **
**---------------------------------------------------------------------------**
**   PROGRAMADOR     |   DATA   |            MOTIVO DA ALTERACAO             **
**---------------------------------------------------------------------------**
** CRISTIANO MACHADO |18/11/2015| Fonte Remodelado, Trabalha com datas dinami**
**                   |          | cas. E esta preparado para receber Plano   **
**                   |          | Mestre..no Futuro                          **
\*---------------------------------------------------------------------------*/
/*
Obseravcoes importantes
  01 = ajustar o ponto de entrada que cria o codigo do produto automatico quando for tipo = 03 nao criar automatico
  02 = nao permitir gerar nota se o produto nao tem NCM.
  03 = alguns produtos que nao tem NCM , VER UM TRATAMENTO.
*/
*******************************************************************************
User Function SG_IMPPREV()
*******************************************************************************

	Private cDirArq			:=	 "C:\TEMP\" + Space(200)

	Private nC4TGer 			:= 0		//| Total Geral de Alterações SC4 |
	Private nC4TInc 			:= 0		//| Total na Inclusões SC4 |
	Private nC4TDel 			:= 0		//| Total de Delecoes SC4 |
	Private nC4TAlt 			:= 0		//| Total de atualizações SC4 |

	Private nH6TGer 			:= 0		//| Total Geral de Alterações SH6 |
	Private nH6TInc 			:= 0		//| Total na Inclusões SH6 |
	Private nH6TDel 			:= 0		//| Total de Delecoes SH6 |
	Private nH6TAlt 			:= 0		//| Total de atualizações SH6 |

	Private aCabec				:= {} //| Array que armazena o Cabecalho |
	Private aDados				:= {} //| Array que armazena os Itens |

	Private bLeArq				:=  {|| OkLeTxt(.F.) }
	Private bAplic				:=  {|| AplicFile(.F.) }



	If TelaInicial() // Apresenta Tela de Inicio

		// Le arquivo texto de entrada
		Processa( bLeArq ,"Importando Arquivo CSV...")

		//| Aplica Arquivo na Tabela SC4 - Previsao de Vendas e SH6 - Plano Mestre
		Processa( bAplic ,"Aplicando Arquivo com Previsao...")


		//| Mensagem Com Resultados ao Final do Processo...SC4 |
		If nC4TGer > 0
			Iw_MsgBox(	'Reg. Novos      : ' 	+ Transform(nC4TInc, '@E 999,999 	'		) + _ENTER +;
				'Reg. Atualizados: ' 	+ Transform(nC4TAlt, '@E 999,999 	'		) + _ENTER +;
				'Reg. Apagados   : ' 	+ Transform(nC4TDel, '@E 999,999 	'		) + _ENTER + _ENTER +;
				'Total Reg. Processados: ' 	+ Transform(nC4TGer, '@E 999,999,999	'	) , 'Resultado do Processamento da Previsão de Vendas... ', "INFO" )
		EndIf

		If nH6TGer > 0
		//| Mensagem Com Resultados ao Final do Processo...SH6 |
			Iw_MsgBox(	'Reg. Novos      : ' 	+ Transform(nH6TInc, '@E 999,999 	') + _ENTER +;
				'Reg. Atualizados: ' 	+ Transform(nH6TAlt, '@E 999,999 	') + _ENTER +;
				'Reg. Apagados   : ' 	+ Transform(nH6TDel, '@E 999,999 	') + _ENTER + _ENTER +;
				'Total Reg. Processados: ' 	+ Transform(nH6TGer, '@E 999,999,999	') , 'Resultado do Processamento do Plano Mestre... ', "INFO" )
		EndIf


	EndIF


Return()
*******************************************************************************
Static Function TelaInicial() // Monta Tela de Inicio...
*******************************************************************************
	Local lOk 			:= .F.

	Local bActOK 	:= {|| lOk := .T. , oWinIni:End() }
	Local bActCl 	:= {|| oWinIni:End() }
	Local cTexto   := "Este Programa Recebe Arquivo Texto no Formato CSV (Separado por ';') Contendo as Quantidades Previstas de Vendas, e as insere na Tabela do Sistema SC4 - Previsao de Vendas."


	oWinIni	:= TDialog():New(001,001,190,380,'Importa Arquivo Com Previsao de Vendas'/* e Plano Mestre...'*/,,,,,CLR_BLACK,CLR_WHITE,,,.T.)

	oMemo		:= TMultiget():Create(oWinIni,{|u|if(Pcount()>0,cTexto:=u,cTexto)},010,010,170,040,,,,,,.T.)
	oSay			:= TSay():Create(oWinIni,{||"Caminho do arquivo : "},055,010,,,,,,.T.,CLR_RED,CLR_WHITE,085,20)
	oGetArq	:= TGet():Create( oWinIni, {|u|If(PCount()==0,cDirArq, cDirArq:=u)},062,010,171,010,"@!",,0,,,.F.,,.T.,,.F.,,.F.,.F.,,.F.,.F.,,'cDirArq',,,, )
	oTButOK 	:= TButton():New( 080, 010, "Confirma"	,oWinIni,bActOK, 40,10,,,.F.,.T.,.F.,,.F.,,,.F. )
	oTButKO 	:= TButton():New( 080, 140, "Cancela"	,oWinIni,bActCl, 40,10,,,.F.,.T.,.F.,,.F.,,,.F. )

	oWinIni:Activate(,,,.T.)

	// Valida se Local e nome do arquivo informado esta Correto...Deve encontrar o arquivo
	If !File(cDirArq) .And. lOk
		Iw_MsgBox('Arquivo nao encontrado!!! '+ Alltrim(cDirArq) +'...','Atencao', ALERT)
		TelaInicial()
		lOk := .F.
	Endif

Return(lOk)
*******************************************************************************
Static Function OkLeTxt //| Funcao para Abertura do arquivo texto e importacao na SB7 |
*******************************************************************************
	Local cLinha := ""
	Local aLinha := {}

	ProcRegua(0)

	FT_FUSE(cDirArq)
	FT_FGOTOP()

	cLinha :=  FT_FREADLN() // Obtem Cabecalho
	aCabec := 	StrTokArr( cLinha, ';')

	FT_FSKIP()

	Do While !FT_FEOF()

		cLinha	 :=	FT_FREADLN()
		aLinha := 	StrTokArr( cLinha, ';')

		Aadd(aDados,aLinha)

		FT_FSKIP()

		IncProc('Processando...Item: '+ aLinha[_PRODUTO] )
	EndDo

	fClose(cDirArq)

Return()

*******************************************************************************
Static Function AplicFile()	//| Aplica Arquivo com Previsoes SC4 e SH6
*******************************************************************************

	Local cProduto 	:= ''
	Local cLocal	:= ''
	Local dData		:= dDatabase
	Local nQtd		:= 0

	ProcRegua(0)

	//DbSelectArea("SH6");DbSetOrder(1)
	DbSelectArea("SC4");DbSetOrder(1)

	For nC := _PRIQTD To Len(aCabec) // Pula Coluna 1 e 2 por Motivo das Quantidades iniciarem na 3 Coluna

		For nI := 1 To Len(aDados)

			cProduto 	:= Padr(Alltrim(aDados[nI][_PRODUTO]),15," ")
			cLocal			:= StrZero(Val(aDados[nI][_LOCAL]),2)
			dData			:= cToD("01"+Substr(aCabec[nC],3))
			nQtd				:= Val	(aDados[nI][nC])

			//| Previsao de Vendas |
			SendSc4( cProduto, cLocal, dData, nQtd ) //| Produto, Local, Data, Qtd |

			//| Plano Mestre |
			//SendSh6( cProduto, cLocal, dData, nQtd ) //| Produto, Local, Data, Qtd |

			IncProc('Processando...Periodo: '+ DToC(dData) )
		Next

	Next

Return()
*******************************************************************************
Static Function SendSc4( cProduto, cLocal, dData, nQtd ) //| Inclui/Altera/deleta Tabela SC4 - Previsao de vendas |
*******************************************************************************
	Local lDel  := .F.
	Local lLock := .F.


	DbSelectArea("SC4")

	//| Verifica se registro já existe...|
	If ( Dbseek(xFilial("SC4")+cProduto+DToS(dData),.F.) )
		lLock := .T.

		RecLock("SC4",.F.)
		If ( nQtd == 0 ) // Tratamento... em caso de atualizacao para Zero deve deletar...
			lDel 		:= .T.
		Else
			nC4TAlt 	+= 1		//| Total de atualizações SC4 |
		EndIf

	ElseIf !( nQtd == 0 )
		RecLock("SC4",.T.)
		lLock 		:= .T.
		nC4TInc 	+=  1		//| Total na Inclusões SC4 |
	EndIf

	//| Apaga Registro....
	If lLock .And. lDel

		DbDelete()
		nC4TDel 		+= 1		//| Total de Delecoes SC4 |

	ElseIf lLock // Inclui ou Atuliza Registro...

		SC4->C4_FILIAL   := xFilial("SC4")
		SC4->C4_PRODUTO  := cProduto
		SC4->C4_DATA     := dData
		SC4->C4_QUANT    := nQtd
		SC4->C4_DOC      := DToS(dDataBase)
		SC4->C4_TIPOPER  := "M"
		SC4->C4_LOTEMIN  := 1
		SC4->C4_LOCAL    := cLocal

	EndIf

	If lLock
		MSUnLock()
		lLock 		:= .F.
		nC4TGer 	+=  1		//| Total Geral de Alterações SC4 |
	EndIf

Return()
*******************************************************************************
Static Function SendSh6( cProduto, cLocal, dData, nQtd ) //| Inclui/Altera/deleta Tabela SH6 - Plano Mestre |
*******************************************************************************
	Local lDel  := .F.
	Local lLock := .F.


	DbSelectArea("SHC")
	//| Verifica se registro já existe...|
	If ( Dbseek(xFilial("SHC")+DToS(dData)+cProduto,.F.) )
		lLock := .T.

		RecLock("SHC",.F.)
		If ( nQtd == 0 ) // Tratamento... em caso de atualizacao para Zero deve deletar...
			lDel 		:= .T.
		Else
			nH6TAlt 	+= 1		//| Total de atualizações SH6 |
		EndIf

	ElseIf !( nQtd == 0 )
		RecLock("SHC",.T.)
		lLock 		:= .T.
		nH6TInc 	+=  1		//| Total na Inclusões SH6 |
	EndIf

	//| Apaga Registro....
	If lLock .And. lDel

		DbDelete()
		nH6TDel 		+= 1		//| Total de Delecoes SH6 |

	ElseIf lLock // Inclui ou Atuliza Registro...

		SHC->HC_FILIAL   := 	xFilial("SHC")
		SHC->HC_PRODUTO  :=  cProduto
		SHC->HC_DATA     :=  DToS(dData)
		SHC->HC_QUANT    :=  nQtd

	EndIf

	If lLock
		MSUnLock()
		lLock 		:= .F.
		nH6TGer 	+=  1		//| Total Geral de Alterações SH6 |
	EndIf

Return()
