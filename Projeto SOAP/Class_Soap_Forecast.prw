#Include 'Totvs.ch'

#Define _ANO_DIAS_  	365
#Define _MES_DIAS_  	30

#Define _ULT_DIA_	 	0
#Define _PRI_DIA_	 	1

#Define _SLD_INI_		'01' 	// Saldo Inicial
#Define _DEV_QTD_		'02' 	// Devolucao
//#Define _PVN_ENT_		'03' 	// Pedidos Nao Entregues
#Define _PRE_VEN_		'03' 	// Previsao de Venda
#Define _QTD_PRO_		'04' 	// Producao
#Define _PRE_NCO_		'05'		// Previsao Nao consumida | Calculo: [03 - 04]
#Define _PVD_FAT_		'06'  	// Pedido Faturado
#Define _PVD_PRE_		'07' 	// Pedido Previsto
#Define _PVA_FIN_		'08'		// Pedido Disponivel | Calculo: [02 + 03] - [03 - 04] - 07
#Define _ORD_PRO_		'09' 	// Ordem Producao
#Define _SAL_FIN_		'10' 	// Saldo Final
#Define _SAL_ATU_		'11' 	// Saldo Atual

#Define _NAC_		 		 01		//| Venda Nacional ... |
#Define _EXP_		 		 02		//| Venda Exportacao ...|

#Define _OPERAC_			"P"		//| Processos ... |
#Define _MEDIAS_			"M"		//| Medias ... |
#Define _CUSCOM_			"C"		//| Custo e Compras ... |

#Define 	_QTD_ 				01			//| Posicao Campo Quantidades no Array AcpoMed[N][N][01]
#Define 	_VLR_				02			//| Posicao Campo Quantidades no Array AcpoMed[N][N][02]

#Define _COM_				01			//| Posicao Campo Ultimo Preco de Compras no Array aCpoComCus
#Define _CUS_				02			//| Posicao Campo Custo Medio  						no Array aCpoComCus

#Define _ENTER_ 			CHR(13) + CHR(10) // | Enter |

*******************************************************************************
User Function SOAPFOREC()
*******************************************************************************

	Local bAction 	:= {|| Executar() }
	Local cTitle		:= "Aguarde..."
	Local cMsg 		:= "Inicializando Processamento..."
	Local lAbort		:= .T.

	Private lMostra := .F.
	Private lChange := .F.

	If Iw_MsgBox("Deseja Gerar o Forecast ?",'Forecast - ' + dToC(dDatabase), "YESNO")

		Processa( bAction, cTitle, cMsg, lAbort )

	EndIf

Return()
*******************************************************************************
Static Function Executar()
*******************************************************************************

	Private oSoap := Class_Soap_Forecast():Novo()

	ProcRegua(0)

	IncProc('Criando arquivo de trabalho...')
	oSoap:CriaTab()

	IncProc('Obtendo Operacoes no periodo...'+_ENTER_+'['+DToC(StoD(oSoap:cDataIni))+'] - ['+DToC(StoD(oSoap:cDataFim))+']' )
	oSoap:VerOperac()

	IncProc(' Obtendo os Valores e Quantidades Medios...')
	oSoap:ObtemMedia()

	IncProc(' Obtendo os Ultimo Preco de Compras e Custo Medio...')
	oSoap:ObtemComCus()

	IncProc('Exportando Tabela Principal para Arquivo CSV...')
	oSoap:ExpCsv()

Return()
*******************************************************************************
Class Class_Soap_Forecast
*******************************************************************************

	//| Propriedades ............................................................

	Data  cAliasTP		//| Alias Tabela Auxiliar Principal	|

	Data  cAliasQO		//| Alias Query Operacoes 	|
	Data  cAliasQM		//| Alias Query Medias 			|
	Data  cAliasCC		//| Alias Query Ultimo Preco Compras e Custo Medio 			|

	Data  cDataIni		//| Data inicial -12 Meses |
	Data  cDataFim		//| Data Final 	+18 Meses |
	Data  cSqlFPro		//| Sql tabela Filtro de Produtos... |

	Data  aDatMedI		//| Array com Datas Iniciais para Cada Periodo das Medias [03 Meses, 06 Meses e 12 Meses ]|
	Data  aDatMedF		//| Array com Datas Finais   para Cada Periodo das Medias [03 Meses, 06 Meses e 12 Meses ]|

	Data  aExpNac		//| Array Filtro Nacional e Exportacao para Query |
	Data  aCpoMed		//| Array Contendo Campos para salvar Medias  |

	//| Metodos ..................................................................

	MethoD Novo() 				CONSTRUCTOR							//| Metodo Construtor inicialisa as Variaveis...|

	MethoD AtuPar(cDataIni, cDataFim)							//| Atualiza os paramentros que devem ser Utilizados...|

	MethoD CriaTab()															//| Cria Tabela de Trabalho Auxiliar Vazia |

	MethoD VerOperac()														//| Executa Query para Obter os Dados de Operacoes.. |

	MethoD ObtemMedia()														//| Obtem as Medias ... |

	MethoD ObtemComCus() 													//| ... |

	MethoD LoadDados(cAliasT , cAliasQ , cQuem )  	//| Transfere Dados da Query para Tabela Auxiliar |

	MethoD ExpCsv(cPath, cFile)										//| Converte a Tabela Auxiliar em Arquivo CSV |


EndClass
*******************************************************************************
Method Novo() Class Class_Soap_Forecast //| Metodo Construtor inicialisa as Variaveis...|
*******************************************************************************
	Local dDtIAux 	:= Lastday( ( dDatabase - ( _MES_DIAS_ * 12 ) ), _PRI_DIA_ ) //| Data inicial Periodo - 12 Meses |
	Local dDtFAux 	:= Lastday( ( dDatabase + ( _MES_DIAS_ * 18 ) ), _ULT_DIA_ ) //| Data Final   Periodo + 18 Meses |

	Local dDtFMan 	:= Lastday( ( dDatabase - ( _MES_DIAS_ / _MES_DIAS_ ) ) , _PRI_DIA_ )	//| Data Final Mes Anterior |

	Local dDt03MI	:=	 Lastday( ( dDatabase - ( _MES_DIAS_ * 03 ) ), _PRI_DIA_ )	//| Data inicial -03 Meses |
	Local dDt06MI	:=	 Lastday( ( dDatabase - ( _MES_DIAS_ * 06 ) ), _PRI_DIA_ )	//| Data inicial -06 Meses |
	Local dDt12MI	:=	 Lastday( ( dDatabase - ( _MES_DIAS_ * 12 ) ), _PRI_DIA_ )	//| Data inicial -12 Meses |


	::cAliasTP 	:= 'TAUX'	//| Alias Tabela Auxiliar 	|
	::cAliasQO		:= 'QOPE'	//| Alias Query Operacoes 	|
	::cAliasQM  	:= 'QMED'	//| Alias Query Medias 			|
	::cAliasCC		:= 'QCEC'	//| Alias Query Ultimo Preco de Comrpas e Custo Medio |


	::cDataIni		:= dTos(dDtIAux)
	::cDataFim  	:= dTos(dDtFAux)

	::aDatMedI		:= {dTos(dDt03MI) , dTos(dDt06MI) , dTos(dDt12MI)}
	::aDatMedF		:= {dTos(dDtFMan) , dTos(dDtFMan) , dTos(dDtFMan)}

	//| Array Filtro Nacional e Exportacao para Query |
	::aExpNac		:= { " AND F2.F2_EST <> 'EX' " , " AND F2.F2_EST = 'EX' " }

	//| Array contendo os campos de Media... |
	::aCpoMed 		:= {{{"QTDM03NA","VLRM03NA"},{"QTDM06NA","VLRM06NA"},{"QTDM12NA","VLRM12NA"}},{{"QTDM03EX","VLRM03EX"},{"QTDM06EX","VLRM06EX"},{"QTDM12EX","VLRM12EX"}}}

	//| Campo caracter que recebe o filtro de produtos... por recno |
	::cSqlFPro		:= DefFilP() //| Define o Filtro de Produtos a serem aplicados |

Return()
*******************************************************************************
Static Function DefFilP()	 //| Define o Filtro de Produtos a serem aplicados |
*******************************************************************************
	Local cSql := ''

	cSQl +=" SELECT R_E_C_N_O_ "
	cSQl +=" FROM SB1010 "
	cSQl +=" WHERE B1_FILIAL = ' '
	cSQl +=" AND B1_GRUPO BETWEEN ' ' AND 'ZZZZZZ' "
	cSQl +=" AND B1_SERIE BETWEEN ' ' AND 'ZZZZZZ' "
	cSQl +=" AND B1_MSBLQL IN('2',' ') "
	cSQl +=" AND B1_TIPO IN('PA','RV') "
	cSQl +=" AND B1_COD <> 'RESERV01001' "
	cSQl +=" AND D_E_L_E_T_ = ' ' "

Return(cSQl)
*******************************************************************************
Method AtuPar(cDataIni, cDataFim) Class Class_Soap_Forecast //| Atualiza os paramentros que devem ser Utilizados...|
*******************************************************************************

	::cDataIni		:= cDataIni	//| Data inicial -12 Meses |
	::cDataFim		:= cDataFim	//| Data Final 	+18 Meses |

Return()
*******************************************************************************
Method CriaTab() Class	Class_Soap_Forecast //| Cria Tabela de Trabalho Auxiliar Vazia |
*******************************************************************************

	Local aStruct 		:= {}
	Local aStruIndex	:= {"SERIE + NOP"}
	Local nPeriodos	:= Round( ( sToD(::cDataFim) - sToD(::cDataIni) ) / _MES_DIAS_,0)

	Local dDataAux		:= ( SToD( ::cDataIni ) - 5 )
	Local oCampo			:= { || dDataAux := ( LastDay(dDataAux,_ULT_DIA_) + 1) , cCampo := 'P_' + Substr(dTos(dDataAux),1,6) }
	Local cCampo			:= ''

	//ALERT("CriaTab")


	Aadd(aStruct , { "SERIE" 				, "C",  60, 0	} ) //| Serie do Produto |
	Aadd(aStruct , { "NOP"					, "C",  02, 0	} ) //| Numero da Operacao |
	Aadd(aStruct , { "OPERACAO"			, "C",  60, 0	} ) //| Nome da Operacao |

	//| Monta os Campos de Periodos... |
	For n := 1 To nPeriodos

		eVal(oCampo)

		Aadd(aStruct , { cCampo 		, "N",  14, 2	} )

	Next

	//| Monta Campos de Medias |
	For nTv := 1 To Len(::aCpoMed) //| For -> Nivel 1 = [ Venda Nacional | Venda Exportacao ]

		For nMe := 1 To Len(::aCpoMed[nTv])	//| For -> Nivel 2 = [3 Meses | 6 Meses | 12 Meses ]|

			For nQV := 1 To Len(::aCpoMed[nTv][nMe]) //| For -> Nivel 3 = [ Quantidade | Valor ]|

				Aadd(aStruct , { ::aCpoMed[nTv][nMe][nQV] 		, "N",  14, 2	} )

			Next
		Next
	Next

	Aadd(aStruct , { "ULTPCMED"			, "N",  14, 2	} ) //| Media Ultimo Preco Serie |
	Aadd(aStruct , { "CUSTOMED"			, "N",  14, 2	} ) //| Media Custo Medio Serie  |

	U_MyFile( aStruct, aStruIndex, Nil, Nil, ::cAliasTP, Nil, .T. )


Return()
*******************************************************************************
Method VerOperac() Class Class_Soap_Forecast //| Executa Query para Obter os Dados.. |
*******************************************************************************

	Local cSql := ''
	//ALERT("VerOperac")

	cSQl +=" 	SELECT SERIE, DATA, SUM(SLD_INI) SLD_INI,  SUM(DEV_QTD) DEV_QTD, SUM(PVN_ENT) PVN_ENT, SUM(PRE_VEN) PRE_VEN, SUM(QTD_PRO) QTD_PRO, SUM(PRE_NCO) PRE_NCO, SUM(PVD_FAT) PVD_FAT, SUM(PVD_PRE) PVD_PRE, SUM(PVA_FIN) PVA_FIN, SUM(ORD_PRO) ORD_PRO, SUM(SAL_FIN) SAL_FIN, SUM(SAL_ATU) SAL_ATU
	cSQl +=" 	FROM ( "

	//|***************************************************************************|
	//||  Query '01' -> Saldo Inicial

	cSQl +=" 	SELECT SUBSTRING(B9_DATA,1,6) DATA, B1_SERIE SERIE ,
	cSQl +=" 	SUM(B9_QINI) SLD_INI, 0  DEV_QTD, 0 	PVN_ENT, 0 PRE_VEN, 0 QTD_PRO, 0 PRE_NCO, 0 PVD_FAT, 0 PVD_PRE, 0 PVA_FIN, 0 ORD_PRO, 0 SAL_FIN, 0 SAL_ATU

	cSQl +=" 	FROM SB9010 B9 "
	cSQl +=" 		INNER JOIN SB1010 B1  "
	cSQl +=" 	 	ON ' '    = B1_FILIAL  "
	cSQl +=" 	 	AND B9_COD = B1_COD "
	cSQl +=" 	WHERE B9_FILIAL = '01' "
	cSQl +=" 	AND   B9_DATA BETWEEN  '"+ ::cDataIni + "' And '"+ ::cDataFim + "' "
	cSQl +=" 	AND   B9.D_E_L_E_T_ = ' '
	cSQl +=" 	AND   B1.R_E_C_N_O_   IN ( "+ ::cSqlFPro +" ) "
	cSQl +=" 	GROUP BY SUBSTRING(B9_DATA,1,6) , B1_SERIE "

	//|***************************************************************************|
	cSQl +=" 	UNION "
	//|***************************************************************************|
	//|| Query '02' -> Devolucao

	cSQl +=" 	SELECT SUBSTRING(D1_EMISSAO,1,6) DATA, B1_SERIE SERIE,
	cSQl +=" 	0 SLD_INI, SUM(D1_QUANT) DEV_QTD, 0 	PVN_ENT, 0 PRE_VEN, 0 QTD_PRO, 0 PRE_NCO, 0 PVD_FAT, 0 PVD_PRE, 0 PVA_FIN, 0 ORD_PRO, 0 SAL_FIN, 0 SAL_ATU

	cSQl +=" 	FROM SD1010 D1 "
	cSQl +=" 	INNER JOIN SF1010 F1 "
	cSQl +=" 	ON F1_FILIAL  = D1_FILIAL "
	cSQl +=" 	AND F1_DOC	    = D1_DOC "
	cSQl +=" 	AND F1_SERIE   = D1_SERIE "
	cSQl +=" 	AND F1_FORNECE = D1_FORNECE "
	cSQl +=" 	AND F1_LOJA    = D1_LOJA "

	cSQl +=" 	INNER JOIN SB1010 B1 "
	cSQl +=" 	ON D1_FILIAL = '01' "
	cSQl +=" 	AND ' ' 	   = B1_FILIAL "
	cSQl +=" 	AND D1_COD	   = B1_COD "

	cSQl +=" 	INNER JOIN SF4010 F4 "
	cSQl +=" 	ON F4_FILIAL     = D1_FILIAL "
	cSQl +=" 	AND F4_CODIGO     = D1_TES "

	cSQl +=" 	WHERE D1_FILIAL     = '01' "
	cSQl +=" 	AND D1.D_E_L_E_T_ = '' "
	cSQl +=" 	AND F1.F1_TIPO    = 'D'  " //-- devolucoes "
	cSQl +=" 	AND F1.F1_EMISSAO BETWEEN  '"+ ::cDataIni + "' And '"+ ::cDataFim + "' "
	cSQl +=" 	AND F1.D_E_L_E_T_ = ' ' "
	cSQl +=" 	AND B1.R_E_C_N_O_  IN ( "+ ::cSqlFPro +" ) "
	cSQl +=" 	AND B1.D_E_L_E_T_ = ' ' "
	cSQl +=" 	AND F4_ESTOQUE    = 'S' "
	cSQl +=" 	AND F4_DUPLIC     = 'S' "
	cSQl +=" 	AND F4.D_E_L_E_T_ = ' ' "
	cSQl +=" 	GROUP BY  SUBSTRING(D1_EMISSAO,1,6), B1_SERIE "

	//|***************************************************************************|
	cSQl +=" 	UNION "
	//|***************************************************************************|
/*	//|| Query '03' -> Pedidos Nao Entregues
	cSQl +=" 	SELECT SUBSTRING(C5_EMISSAO,1,6) DATA, B1_SERIE SERIE, "
	cSQl +=" 	0 SLD_INI, 0 DEV_QTD, SUM(C6_QTDVEN) PVN_ENT, 0 PRE_VEN, 0 QTD_PRO, 0 PRE_NCO, 0 PVD_FAT, 0 PVD_PRE, 0 PVA_FIN, 0 ORD_PRO, 0 SAL_FIN, 0 SAL_ATU

	cSQl +=" 	FROM SC5010 C5 "
	cSQl +=" 	INNER JOIN SC6010 C6 "
	cSQl +=" 	ON C5_FILIAL = C6_FILIAL "
	cSQl +=" 	AND C5_NUM    = C6_NUM "
	cSQl +=" 	INNER JOIN SB1010 B1 "
	cSQl +=" 	ON ' '       = B1_FILIAL "
	cSQl +=" 	AND C6_PRODUTO = B1_COD "
	cSQl +=" 	INNER JOIN SF4010 F4 "
	cSQl +=" 	ON F4_FILIAL  = C6_FILIAL "
	cSQl +=" 	AND F4_CODIGO  = C6_TES "

	cSQl +=" 	WHERE C5_FILIAL = '01' "
	cSQl +=" 	AND C5_EMISSAO BETWEEN  '"+ ::cDataIni + "' And  '"+ ::dDtFMan + "' " // Sî PEGA PEDIDOS EM ATRASO
	cSQl +=" 	AND C5.D_E_L_E_T_ = ' ' "
	//cSQl +=" 	AND C6_QTDVEN = C6_QTDENT "
	cSQl +=" 	AND C6_QTDVEN > C6_QTDENT "
	cSQl +=" 	AND C6.D_E_L_E_T_ = ' ' "
	cSQl +=" 	AND F4_ESTOQUE    = 'S' "
	cSQl +=" 	AND F4_DUPLIC     = 'S' "
	cSQl +=" 	AND F4.D_E_L_E_T_ = ' ' "
	cSQl +=" 	AND B1.R_E_C_N_O_  IN ( "+ ::cSqlFPro +" ) "
	cSQl +=" 	GROUP BY SUBSTRING(C5_EMISSAO,1,6), B1_SERIE "

	//|***************************************************************************|
	cSQl +="	UNION "
	//|***************************************************************************|
*/	//|| Query '03' -> Previsao de Venda

	cSQl +=" 	SELECT SUBSTRING(C4_DATA,1,6) DATA, B1_SERIE SERIE, "
	cSQl +=" 	0 SLD_INI, 0 DEV_QTD, 0 PVN_ENT, SUM(C4_QUANT) PRE_VEN, 0 QTD_PRO, 0 PRE_NCO, 0 PVD_FAT, 0 PVD_PRE, 0 PVA_FIN, 0 ORD_PRO, 0 SAL_FIN, 0 SAL_ATU
	cSQl +=" 	FROM SC4010 C4 "
	cSQl +=" 	INNER JOIN SB1010 B1 "
	cSQl +=" 	ON C4_FILIAL  = '01' "
	cSQl +=" 	AND C4_PRODUTO = B1_COD "
	cSQl +=" 	AND ' ' 	    = B1_FILIAL "

	cSQl +=" 	WHERE C4.C4_FILIAL = '01' "
	cSQl +=" 	AND   C4.C4_DATA BETWEEN '" + ::cDataIni + "' And '"+ ::cDataFim + "' "  //--// -12m +12m
	cSQl +=" 	AND   C4.D_E_L_E_T_ = ' ' "
	cSQl +=" 	AND   B1.D_E_L_E_T_ = ' ' "
	cSQl +=" 	AND   B1.R_E_C_N_O_ IN ( "+ ::cSqlFPro +" ) "
	cSQl +=" 	GROUP BY SUBSTRING(C4_DATA,1,6) , B1_SERIE "

	//|***************************************************************************|
	cSQl +="	UNION "
	//|***************************************************************************|
	//|| Query '04' -> Producao
	cSQl +=" 	SELECT SUBSTRING(H6_DTAPONT,1,6) DATA, B1_SERIE SERIE, "
	cSQl +=" 	0 SLD_INI, 0 DEV_QTD, 0 PVN_ENT, 0 PRE_VEN, SUM(H6_QTDPROD) QTD_PRO, 0 PRE_NCO, 0 PVD_FAT, 0 PVD_PRE, 0 PVA_FIN, 0 ORD_PRO, 0 SAL_FIN, 0 SAL_ATU

	cSQl +=" 	FROM SH6010 H6 "
	cSQl +=" 	INNER JOIN (	SELECT G2_FILIAL, G2_PRODUTO, B1_SERIE, MAX(G2_OPERAC) G2_OPERAC "
	cSQl +=" 	FROM SG2010 G2 "
	cSQl +=" 	INNER JOIN SB1010 B1 "
	cSQl +=" 	ON ' ' 	    = B1_FILIAL "
	cSQl +=" 	AND G2_FILIAL  = '01' "
	cSQl +=" 	AND G2_PRODUTO = B1_COD "

	cSQl +=" 	WHERE G2_FILIAL = '01' "
	cSQl +=" 	AND G2.D_E_L_E_T_ = ' ' "
	cSQl +=" 	AND B1.R_E_C_N_O_   IN ( "+ ::cSqlFPro +" ) "
	cSQl +=" 	GROUP BY G2_FILIAL, G2_PRODUTO, B1_SERIE ) SG2 "

	cSQl +=" 	ON SG2.G2_FILIAL	= H6.H6_FILIAL "
	cSQl +=" 	AND SG2.G2_OPERAC	= H6.H6_OPERAC "
	cSQl +=" 	AND SG2.G2_PRODUTO = H6.H6_PRODUTO "

	cSQl +=" 	WHERE H6_FILIAL = '01' "
	cSQl +=" 	AND H6_DTAPONT BETWEEN  '" + ::cDataIni + "'  And  '"+ ::cDataFim + "' "
	cSQl +=" 	AND H6.D_E_L_E_T_ = ' ' "
	cSQl +=" 	GROUP BY SUBSTRING(H6_DTAPONT,1,6), B1_SERIE "

	//|***************************************************************************|
	//	cSQl +="	UNION "
	//|***************************************************************************|

	//|| Query '05' -> Previsao Nao consumida | Calculo: [03 - 04]
	//	cSQl +=" 	0 SLD_INI, 0 DEV_QTD, 0 PVN_ENT, 0 PRE_VEN, 0 QTD_PRO, 0 PRE_NCO, 0 PVD_FAT, 0 PVD_PRE, 0 PVA_FIN, 0 ORD_PRO, 0 SAL_FIN, 0 SAL_ATU

	//|***************************************************************************|
	cSQl +=" UNION "
	//|***************************************************************************|

	//|| Query '06' -> Pedido Faturado
	cSQl +=" 	SELECT SUBSTRING(D2_EMISSAO,1,6) DATA, B1_SERIE SERIE, "
	cSQl +=" 	0 SLD_INI, 0 DEV_QTD, 0 PVN_ENT, 0 PRE_VEN, 0 QTD_PRO, 0 PRE_NCO, SUM(D2_QUANT) PVD_FAT, 0 PVD_PRE, 0 PVA_FIN, 0 ORD_PRO, 0 SAL_FIN, 0 SAL_ATU

	cSQl +=" 	FROM SD2010 D2 "
	cSQl +=" 	INNER JOIN SF2010 F2 "
	cSQl +=" 	ON F2_FILIAL = D2_FILIAL "
	cSQl +=" 	AND F2_DOC	   = D2_DOC "
	cSQl +=" 	AND F2_SERIE  = D2_SERIE "
	cSQl +=" 	AND F2_CLIENT = D2_CLIENTE "
	cSQl +=" 	AND F2_LOJA   = D2_LOJA "

	cSQl +=" 	INNER JOIN SB1010 B1 "
	cSQl +=" 	ON D2_FILIAL = '01' "
	cSQl +=" 	AND ' ' 	   = B1_FILIAL "
	cSQl +=" 	AND D2_COD	   = B1_COD "

	cSQl +=" 	INNER JOIN SF4010 F4 "
	cSQl +=" 	ON F4_FILIAL     = D2_FILIAL "
	cSQl +=" 	AND F4_CODIGO     = D2_TES "

	cSQl +=" 	WHERE D2_FILIAL     = '01' "
	cSQl +=" 	AND D2.D_E_L_E_T_ = '' "
	cSQl +=" 	AND F2.F2_TIPO    = 'N' "
	cSQl +=" 	AND F2.F2_EMISSAO BETWEEN  '"+ ::cDataIni + "' And '"+ ::cDataFim + "' "
	cSQl +=" 	AND F2.D_E_L_E_T_ = ' ' "
	cSQl +=" 	AND B1.R_E_C_N_O_  IN ( "+ ::cSqlFPro +" ) "
	cSQl +=" 	AND B1.D_E_L_E_T_ = ' ' "
	cSQl +=" 	AND F4_ESTOQUE    = 'S' "
	cSQl +=" 	AND F4_DUPLIC     = 'S' "
	cSQl +=" 	AND F4.D_E_L_E_T_ = ' ' "
	cSQl +=" 	GROUP BY SUBSTRING(D2_EMISSAO,1,6), B1_SERIE "

	//|***************************************************************************|
	cSQl +=" 	UNION "
	//|***************************************************************************|
	//|| Query '07' -> Pedido Previsto
	cSQl +=" 	SELECT SUBSTRING(C5_EMISSAO,1,6) DATA, B1_SERIE SERIE, "
	cSQl +=" 	0 SLD_INI, 0 DEV_QTD, 0 PVN_ENT, 0 PRE_VEN, 0 QTD_PRO, 0 PRE_NCO, 0 PVD_FAT, SUM(C6_QTDVEN) PVD_PRE, 0 PVA_FIN, 0 ORD_PRO, 0 SAL_FIN, 0 SAL_ATU

	cSQl +=" 	FROM SC5010 C5 "
	cSQl +=" 	INNER JOIN SC6010 C6 "
	cSQl +=" 	ON C5_FILIAL = C6_FILIAL "
	cSQl +=" 	AND C5_NUM    = C6_NUM "
	cSQl +=" 	INNER JOIN SB1010 B1 "
	cSQl +=" 	ON ' '       = B1_FILIAL "
	cSQl +=" 	AND C6_PRODUTO = B1_COD "
	cSQl +=" 	INNER JOIN SF4010 F4 "
	cSQl +=" 	ON F4_FILIAL  = C6_FILIAL "
	cSQl +=" 	AND F4_CODIGO  = C6_TES "

	cSQl +=" 	WHERE C5_FILIAL = '01' "
	cSQl +=" 	AND C5_EMISSAO BETWEEN  '"+ ::cDataIni + "' And  '"+ ::cDataFim + "' " // SO PEGA PEDIDOS EM ABERTO PREVISTO
	cSQl +=" 	AND C5.D_E_L_E_T_ = ' ' "
	cSQl +=" 	AND C6_QTDVEN > C6_QTDENT "
	cSQl +=" 	AND C6.D_E_L_E_T_ = ' ' "
	cSQl +=" 	AND F4_ESTOQUE    = 'S' "
	cSQl +=" 	AND F4_DUPLIC     = 'S' "
	cSQl +=" 	AND F4.D_E_L_E_T_ = ' ' "
	cSQl +=" 	AND B1.R_E_C_N_O_  IN ( "+ ::cSqlFPro +" ) "
	cSQl +=" 	GROUP BY SUBSTRING(C5_EMISSAO,1,6), B1_SERIE "

	//|***************************************************************************|
	//	cSQl +="	UNION "
	//|***************************************************************************|

	//|| Query '08' -> Pedido Disponivel | Calculo: [02 + 03] - [03 - 04] - 07
	//	cSQl +=" 	0 SLD_INI, 0 DEV_QTD, 0 PVN_ENT, 0 PRE_VEN, 0 QTD_PRO, 0 PRE_NCO, 0 PVD_FAT, 0 PVD_PRE, 0 PVA_FIN, 0 ORD_PRO, 0 SAL_FIN, 0 SAL_ATU

	//|***************************************************************************|
	cSQl +=" UNION "
	//|***************************************************************************|
	//|| Query '09' -> Ordem Producao

	cSQl +=" 	SELECT SUBSTRING(C2_DATRF,1,6) DATA, B1_SERIE SERIE, "
	cSQl +=" 	0 SLD_INI, 0 DEV_QTD, 0 PVN_ENT, 0 PRE_VEN, 0 QTD_PRO, 0 PRE_NCO, 0 PVD_FAT, 0 PVD_PRE, 0 PVA_FIN, SUM(C2_QUJE) ORD_PRO, 0 SAL_FIN, 0 SAL_ATU "

	cSQl +=" 	FROM SC2010 C2 inner Join SB1010 B1 "
	cSQl +=" 	ON	B1_FILIAL		= '' "
	cSQl +=" 	AND C2.C2_FILIAL	= '01' "
	cSQl +=" 	AND C2.C2_PRODUTO	= B1.B1_COD "

	cSQl +=" 	WHERE C2.C2_FILIAL  = '01' "
	cSQl +=" 	AND   C2.C2_DATRF   BETWEEN  '"+ ::cDataIni + "' And '"+ ::cDataFim + "' "  --// -12m +18m
	cSQl +=" 	AND   C2.D_E_L_E_T_ = '' "
	cSQl +=" 	AND   B1.R_E_C_N_O_  IN ( "+ ::cSqlFPro +" ) "
	cSQl +=" 	GROUP BY  C2_FILIAL,  SUBSTRING(C2_DATRF,1,6), B1_SERIE "


	//|***************************************************************************|
	//	cSQl +="	UNION "
	//|***************************************************************************|

	//|| Query '10' -> Saldo Final
	//	cSQl +=" 	0 SLD_INI, 0 DEV_QTD, 0 PVN_ENT, 0 PRE_VEN, 0 QTD_PRO, 0 PRE_NCO, 0 PVD_FAT, 0 PVD_PRE, 0 PVA_FIN, 0 ORD_PRO, 0 SAL_FIN, 0 SAL_ATU

	//|***************************************************************************|
	cSQl +=" 	UNION "
	//|***************************************************************************|
	//|| Query '11' -> Saldo Atual
	cSQl +=" 	SELECT "+Substr(DToS(dDataBase),1,6)+" DATA, B1_SERIE SERIE, "//,    0  PRE_VEN,  0  QTD_PRO,  0  ORD_PRO,  0  PVD_FAT,  0  DEV_QTD,   0  PVN_ENT,  0  PVD_FAT,  0  PV_ENTRE_EANT,   0  SLD_INI, SUM(B2_QATU) SAL_ATU "
	cSQl +=" 	0 SLD_INI, 0 DEV_QTD, 0 PVN_ENT, 0 PRE_VEN, 0 QTD_PRO, 0 PRE_NCO, 0 PVD_FAT, 0 PVD_PRE, 0 PVA_FIN, 0 ORD_PRO, 0 SAL_FIN,   SUM(B2_QATU-B2_QPEDVEN-B2_RESERVA) SAL_ATU  "
	cSQl +=" 	FROM SB2010 B2 "
	cSQl +=" 	INNER JOIN SB1010 B1 "
	cSQl +=" 	ON ' '    = B1_FILIAL "
	cSQl +=" 	AND B2_COD = B1_COD "

	cSQl +=" 	WHERE B2_FILIAL = '01' "
	cSQl +=" 	AND   B2.D_E_L_E_T_ = ' ' "
	cSQl +=" 	AND   B1.R_E_C_N_O_   IN ( "+ ::cSqlFPro +" ) "

	cSQl +=" 	GROUP BY B1_SERIE "

	//|***************************************************************************|

	cSQl +=" 	) TODAS "

	cSQl +=" 	GROUP BY SERIE, DATA"

	cSQl +=" 	ORDER BY SERIE, DATA"

	U_ExecMySql( cSql , ::cAliasQO , 'Q'	, lMostra, lChange )

	//| Transfere Dados da Query para Tabela Auxiliar |
	::LoadDados(::cAliasTP , ::cAliasQO , _OPERAC_ , NIL )


Return()
*******************************************************************************
MethoD ObtemMedia() 	Class	Class_Soap_Forecast 	//| Obtem as Medias ... |
*******************************************************************************
	//ALERT("ObtemMedia")
	//| Monta Campos de Medias |
	For nTv := 1 To Len(::aCpoMed) //| For -> Nivel 1 = [ Venda Nacional | Venda Exportacao ]

		For nMe := 1 To Len(::aCpoMed[nTv])	//| For -> Nivel 2 = [3 Meses | 6 Meses | 12 Meses ]|

			//| Executa Query para Obter as Medias conforme os parametros |
			QueryMedia( ::aExpNac[nTv] , ::aDatMedI[nMe] , ::aDatMedF[nMe], ::cSqlFPro )

 			//| Transfere Dados da Query para Tabela Auxiliar Principal |
			::LoadDados(::cAliasTP , ::cAliasQM , _MEDIAS_, ::aCpoMed[nTv][nMe] )

		Next

	Next

Return()
*******************************************************************************
Static Function  QueryMedia( cTpV , cDataI , cDataF, cSqlFPro )	 	//| Query Default para Medias.... |
*******************************************************************************
	Local cAliasQ := 'QMED'

	If Select(	cAliasQ ) == 0
		DbCloseArea()
	EndIf

	//ALERT("QueryMedia")

	cSQl := ""

	cSQl += " 	SELECT	SUBSTRING(D2_EMISSAO,1,6) DATA, B1_SERIE SERIE, "
	cSQl += " 			ROUND(( SUM(D2_QUANT)	  / COUNT(1) ), 2 )	  Q_MEDIA, "
	cSQl += " 			ROUND(( SUM(D2_PRCVEN)	/ COUNT(1) ), 2 )  V_MEDIA	 "

	cSQl += " 	FROM SD2010 D2  "
	cSQl += " 	INNER JOIN SF2010 F2 "
	cSQl += " 	ON F2_FILIAL = D2_FILIAL  "
	cSQl += " 	AND F2_DOC	   = D2_DOC  "
	cSQl += " 	AND F2_SERIE  = D2_SERIE  "
	cSQl +=" 	AND F2_CLIENT = D2_CLIENTE "
	cSQl +=" 	AND F2_LOJA   = D2_LOJA "

	cSQl +=" 	INNER JOIN SB1010 B1  "
	cSQl +=" 	ON D2_FILIAL = '01'  "
	cSQl +=" 	AND ' ' 	   = B1_FILIAL  "
	cSQl +=" 	AND D2_COD	   = B1_COD  "

	cSQl +=" 	INNER JOIN SF4010 F4  "
	cSQl +=" 	ON F4_FILIAL     = D2_FILIAL  "
	cSQl +=" 	AND F4_CODIGO    = D2_TES  "

	cSQl +=" 	WHERE D2_FILIAL     = '01'  "
	cSQl +=" 	AND D2.D_E_L_E_T_ = ''  "
	cSQl +=" 	AND F2.F2_TIPO    = 'N'  "

	cSQl +=" 	AND F2.F2_EMISSAO BETWEEN  '"+ cDataI + "' And '"+ cDataF + "' "

	cSQl +=  	cTpV ///| Tipo de Venda -> Nacional | Exportacao

	cSQl +=" 	AND F2.D_E_L_E_T_ = ' '
	cSQl +=" 	AND B1.R_E_C_N_O_  IN ( "+ cSqlFPro +" ) "
	cSQl +=" 	AND B1.D_E_L_E_T_ = ' '  "
	cSQl +=" 	AND F4_ESTOQUE    = 'S'  "
	cSQl +=" 	AND F4_DUPLIC     = 'S'  "
	cSQl +=" 	AND F4.D_E_L_E_T_ = ' ' "

	cSQl +=" 	GROUP BY SUBSTRING(D2_EMISSAO,1,6), B1_SERIE "

	U_ExecMySql( cSql , cAliasQ , 'Q', lMostra, lChange )

Return()
*******************************************************************************
MethoD ObtemComCus( cTpV , cDataI , cDataF, cSqlFPro )	 Class	Class_Soap_Forecast  //| Query Default para Ultimo Preco de Compra e Custos Medio.... |
*******************************************************************************
//ALERT("ObtemComCus")
	cSQl := ""

	cSQl += " SELECT DATA, SERIE, SUM(MED_PRCCOM) MED_PRCCOM , SUM(MED_CUSMED) MED_CUSMED "
	cSQl += " FROM ( "

	cSQl += " 				 SELECT '201508' DATA, B1_SERIE SERIE , ROUND(SUM(B1_UPRC) / COUNT(1), 2) MED_PRCCOM  , SUM(0) MED_CUSMED "
	cSQl += " 				 FROM SB1010 B1 "
	cSQl += " 				 WHERE B1.R_E_C_N_O_  IN ( "+ ::cSqlFPro +" ) "
	cSQl += " 				 GROUP BY B1_SERIE "

	cSQl += "  UNION "

	cSQl += " 				SELECT '201508' DATA, B1_SERIE SERIE, SUM(0) MED_PRCCOM, ROUND( SUM(B2_CM1)/COUNT(1),2) MED_CUSMED "
	cSQl += " 				FROM SB2010 B2  	INNER JOIN SB1010 B1  	ON ' '    = B1_FILIAL  	AND B2_COD = B1_COD "
	cSQl += " 				WHERE B2_FILIAL = '01'  	AND   B2.D_E_L_E_T_ = ' '  	AND   B1.R_E_C_N_O_   IN ( "+ ::cSqlFPro +" ) "
	cSQl += " 				GROUP BY B1_SERIE "

	cSQl += " 			) PRCCUS "
	cSQl += " GROUP BY DATA, SERIE "
	cSQl += " 	ORDER BY DATA, SERIE "


	U_ExecMySql( cSql , ::cAliasCC , 'Q', lMostra, lChange )

	::LoadDados(::cAliasTP , ::cAliasCC , _CUSCOM_, { "ULTPCMED" , "CUSTOMED" } )

Return()
*******************************************************************************
MethoD LoadDados(cAliasT , cAliasQ , cQuem, aCampos )  Class	Class_Soap_Forecast 	//| Transfere Dados da Query para Tabela Auxiliar |
*******************************************************************************
	//ALERT("LoadDados")

	//ALERT("cAliasT "+cAliasT)
	//ALERT("cAliasQ "+cAliasQ)
	//ALERT("cQuem "+cQuem)



	DbSelectArea( cAliasT );DbGoTop()
	DbSelectArea( cAliasQ );DbGoTop()

	While !EOF()     //cSerie == ((cAliasQ)->SERIE)

		If cQuem == _OPERAC_      //| cQuem -> Quem chamou a Funcao LoadDados ... |
			//ALERT("PRE- > SaveProcesso")
			//ALERT("(cAliasQ)->DATA : " + cValToChar((cAliasQ)->DATA) )
			SaveProcesso(cAliasQ, "P_" + cValToChar((cAliasQ)->DATA) )
		EndIf

		If cQuem == _MEDIAS_  //| cQuem -> Quem chamou a Funcao LoadDados ... |
			SaveMedia(cAliasQ, aCampos)
		EndIf

		If cQuem == _CUSCOM_ //| cQuem -> Quem chamou a Funcao LoadDados ... |
			SaveCusCom(cAliasQ, aCampos)
		EndIf

		DbSelectarea(cAliasQ)
		DbSkip()

	EndDo
	DbSelectArea( cAliasQ )
	DbClosearea()

Return()
*******************************************************************************
Static Function SaveProcesso(cAliasQ, cCpoData) //| Prepara os Processos para Salvar na Tabela Principal |
*******************************************************************************

	Local 	bPRE_NCO_	:= {|| ( nX := (cAliasQ)->PRE_VEN  - (cAliasQ)->QTD_PRO ) , IiF(nX<0,0,nX) }
	Local	bPVA_FIN_	:= {|| ( nY := (cAliasQ)->DEV_QTD	 + (cAliasQ)->PRE_VEN - eVal(bPRE_NCO_) - (cAliasQ)->PVD_FAT )  , IiF(nY<0,0,nY) }
	Local	bSAL_FIN_	:= {|| ( nZ := (cAliasQ)->SLD_INI	 + (cAliasQ)->DEV_QTD + (cAliasQ)->QTD_PRO - (cAliasQ)->PVD_FAT - (cAliasQ)->PVD_PRE )	, IiF(nZ<0,0,nZ) }

//ALERT("SaveProcesso")
		//#Define _SLD_INI_		'01' 	// Saldo Inicial
	SalvaReg(_SLD_INI_		, 	(cAliasQ)->SLD_INI			,	'SALDO INICIAL'				, cCpoData )

		//#Define _DEV_QTD_		'02' 	// Devolucao
	SalvaReg(_DEV_QTD_		, 	(cAliasQ)->DEV_QTD			,	'DEVOLUCAO'						, cCpoData )

		//#Define _PVN_ENT_		'03' 	// Pedidos
	//SalvaReg(_PVN_ENT_		, 	(cAliasQ)->PVN_ENT			,	'PEDIDO NAO ENTREGUE' 	, cCpoData )

		//#Define _PRE_VEN_		'03' 	// Previsao de Venda
	SalvaReg(_PRE_VEN_		, 	(cAliasQ)->PRE_VEN			, 	'PREVISAO DE VENDA'			, cCpoData )

		//#Define _QTD_PRO_		'04' 	// Producao
	SalvaReg(_QTD_PRO_		, 	(cAliasQ)->QTD_PRO			,	'PRODUCAO'							, cCpoData )

		//#Define _PRE_NCO_		'05'	// Previsao Nao consumida | Calculo: [03 - 04]
	SalvaReg(_PRE_NCO_		,	eVal(bPRE_NCO_)				,	'PREVISAO DISPONIVEL'		, cCpoData )

		//#Define _PVD_FAT_		'06'  	// Pedido Faturado
	SalvaReg(_PVD_FAT_		,	(cAliasQ)->PVD_FAT			,	'PEDIDO FATURADO '			, cCpoData )

		//#Define _PVD_PRE_		'07' 	// Pedido Previsto
	SalvaReg(_PVD_PRE_		, 	(cAliasQ)->PVD_PRE			,	'PEDIDO PREVISTO'			, cCpoData )

		//#Define _PVA_FIN_		'08'	// Pedido Atrasado | Calculo: [02 + 03] - [03 - 04] - 07
	SalvaReg(_PVA_FIN_		,	eVal(bPVA_FIN_)				,	'PEDIDO DISPONIVEL'			, cCpoData )

		//#Define _ORD_PRO_		'09' 	// Ordem Producao
	SalvaReg(_ORD_PRO_		,	(cAliasQ)->ORD_PRO 			,	'ORDEM DE PRODUCAO'			, cCpoData )

		//#Define _SAL_FIN_		'10' 	// Saldo Final
	SalvaReg(_SAL_FIN_		,	eVal(bSAL_FIN_)	 			,	'SALDO FINAL'					, cCpoData )

		//#Define _SAL_ATU_		'11' 	// Saldo Atual
	SalvaReg(_SAL_ATU_		, 	(cAliasQ)->SAL_ATU			,	'SALDO ATUAL'					, cCpoData )


Return()
*******************************************************************************
Static Function SaveMedia(cAliasQ , aCampoMedia )//| Prepara as Medias para Salvar na Tabela Principal |
*******************************************************************************

	Local cCpoMV 	:= aCampoMedia[_QTD_] 	//| Campo Quantidade da Tabela Principal
	Local cCpoMQ 	:= aCampoMedia[_VLR_]	//| Campo Valor da Tabela Principal |
//ALERT("SaveMedia")
	//| Foi definido que as medias Ficariam salvas ao lado do Faturamento na mesma linha
	//| por isso estou utilizando o nOp = _PVD_FAT_

	SalvaReg(_PVD_FAT_, (cAliasQ)->Q_MEDIA, 'PEDIDO FATURADO '	, cCpoMV  )

	SalvaReg(_PVD_FAT_, (cAliasQ)->V_MEDIA, 'PEDIDO FATURADO '	, cCpoMQ  )


Return()

*******************************************************************************
Static Function SaveCusCom(cAliasQ, aCpoComCus)
*******************************************************************************

	Local cCpoCom 	:= aCpoComCus[_COM_] 	//| Campo Quantidade da Tabela Principal
	Local cCpoCus 	:= aCpoComCus[_CUS_]	//| Campo Valor da Tabela Principal |
//ALERT("SaveCusCom")
	//| Foi definido que as medias Ficariam salvas ao lado do Faturamento na mesma linha
	//| por isso estou utilizando o nOp = _PVD_FAT_

	SalvaReg(_SAL_ATU_, (cAliasQ)->MED_PRCCOM, 'SALDO ATUAL'	, cCpoCom )

	SalvaReg(_SAL_ATU_, (cAliasQ)->MED_CUSMED, 'SALDO ATUAL'	, cCpoCus )


Return()
*******************************************************************************
Static Function SalvaReg(cNop, nQtd, cOperacao, cCampo  )// Salva um Registro de Query em Tabela Principal
*******************************************************************************
	Local cAliasQ		:= ""
	Local cAliasT		:= 'TAUX' /// ::cAliasTP
//ALERT("SalvaReg")
	If Alltrim(ProcName(1)) == "SAVEPROCESSO"
		cAliasQ := 'QOPE'
	ElseIf  Alltrim(ProcName(1)) == "SAVEMEDIA"
		cAliasQ := 'QMED'
	ElseIf  Alltrim(ProcName(1)) == "SAVECUSCOM"
		cAliasQ := 'QCEC'

	EndIf

	// Caso o Parametro esteja vazio...
	If Empty(cCampo)
		Return()
	EndIF

	DbSelectArea(cAliasT)
	If DbSeek( (cAliasQ)->SERIE + cNop, .F. )

		RecLock(cAliasT, .F.)

	Else
		RecLock(cAliasT, .T.)

		(cAliasT)->SERIE			:= (cAliasQ)->SERIE
		(cAliasT)->NOP				:= cNop
		(cAliasT)->OPERACAO		:= cOperacao

	EndIf

	(cAliasT)->&cCampo 			+= nQtd

	MsUnlock()

Return()
*******************************************************************************
MethoD ExpCsv()	 Class Class_Soap_Forecast
*******************************************************************************
	Local cAliasT 			:= ::cAliasTP
	Local lHeader			:= .T.
	Local aArrayTab		:= {}
	Local cNameFile		:= 'forecast'
	Local cExt					:= '.csv'
	Local cDelimitador	:= ';'
	Local cFileOK			:= ''

	Local cFileFull		:= "C:\TEMP\" + cNameFile + cExt

	aArrayTab 	:= U_TabToArray( cAliasT, lHeader )

	cFileOK		:= U_ArrayToFCsv( aArrayTab , cNameFile, cDelimitador )

	If File("C:\" + cNameFile + cExt)
		FErase("C:\" + cNameFile + cExt )
	EndIf


	If __CopyFIle('\'+cFileOK, cFileFull  )
		ShellExecute("Open", cFileFull, "", "" , 1 )

	Else
		Alert('Arquivo criado em...Servidor\' + cFileOK)

	EndIf


Return()