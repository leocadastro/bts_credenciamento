<%
Erro = ""

	' Limpando campos Hidden
	id_edicao		= limpar_texto(Request("id_edicao"))
	id_idioma		= limpar_texto(Request("id_idioma"))
	id_tipo			= limpar_texto(Request("id_tipo"))
	origem_cnpj		= limpar_texto(Request("origem_cnpj"))
	origem_cpf		= limpar_texto(Request("origem_cpf"))
	id_empresa		= limpar_texto(Request("id_empresa"))
	id_visitante	= limpar_texto(Request("id_visitante"))

	If Len(Trim(id_empresa)) > 0 Then Novo_ID_Empresa = id_empresa

	If id_empresa = "UNDEFINED" Then id_empresa = ""
	If id_visitante = "UNDEFINED" Then id_visitante = ""

	Select Case ID_Formulario
		Case 1, 2, 5
			If Len(Trim(Request("frmCNPJ"))) = 0 AND id_idoma = 1 Then response.Redirect("/?erro=1")
	End Select
	
	' Limpando o CNPJ
	CNPJ 			= limpar_texto(Request("frmCNPJ"))
	CNPJ 			= Replace(CNPJ,".","")
	CNPJ 			= Replace(CNPJ,"-","")
	CNPJ 			= Replace(CNPJ,"/","")
	If Len(CNPJ) <> 14 AND id_idioma = 1 Then
		Erro = Erro & "{ ""Erro"", ""01 - CNPJ Incorreto"" },"
	End If
	
	Razao			= limpar_texto(Request("frmRazao"))			' Value: BTS INFORMA
	Select Case ID_Formulario
		Case 1
			Fantasia 		= limpar_texto(Request("frmFantasia"))		' Value: BTS INFORMA
		Case 2
			Sigla 			= limpar_texto(Request("frmSigla"))			' Value: BTS INFORMA
			Resp 			= limpar_texto(Request("frmResp"))			' Value: BTS INFORMA
		Case 5
			Fantasia 		= limpar_texto(Request("frmFantasia"))		' Value: BTS INFORMA
			Resp 			= limpar_texto(Request("frmResp"))			' Value: BTS INFORMA
			Senha			= limpar_texto(Request("frmSenha"))			' Value: BTS INFORMA
	End Select

	OptRamo			= limpar_texto(Request("frmRamo"))					' Value: 9
	OptRamoCompl	= limpar_texto(Request("frmOptRamoComplemento"))	' Value: 9
	
	'OptRamoOutros	= limpar_texto(Request("frmOptRamoOutros"))			' Value: ''
	'Atividade 		= limpar_texto(Request("frmAtividade"))
	'AtividadeOutros = limpar_texto(Request("frmAtividadeOutros"))
	
	'PriProdut 		= limpar_texto(Request("frmPriProdut"))		' Value: RRR
	produtos_inserir= Request("produtos_inserir")	' Value: RRR

	'response.write("produtos " & produtos_inserir)
	'response.end()

	' Limpando o CEP
	CEP				= limpar_texto(Request("frmCEP"))			' Value: 04583-110
	CEP				= Replace(CEP,"-","")						' Value: 04583-110
	If Len(CEP) <> 8 AND id_idioma = 1 Then
		Erro = Erro & "{ ""Erro"", ""03 - CEP Incorreto"" },"
	End If

	Endereco	 	= limpar_texto(Request("frmEndereco"))		' Value: AV. DR. CHUCRI ZAIDAN
	Numero	 		= limpar_texto(Request("frmNumero"))		' Value: 80
	Complemento		= limpar_texto(Request("frmComplemento")) 	' Value: 2º ANDAR
	Bairro			= limpar_texto(Request("frmBairro")) 		' Value: MORUMBI
	Cidade 			= limpar_texto(Request("frmCidade"))		' Value: SÃO PAULO
	Estado 			= limpar_texto(Request("frmEstado"))		' Value: 26
	Pais			= limpar_texto(Request("frmPais")) 			' Value: 23

	' Limpando o Site
	Site 			= Lcase(limpar_texto(Request("frmSite")))	' Value: www.btsinforma.com.br
	Site 			= Replace(Site,"http://","")				' Value: www.btsinforma.com.br
	Site 			= Replace(Site,"/","")						' Value: www.btsinforma.com.br
	
	QtdFunc 		= limpar_texto(Request("frmQtdFunc"))		' Value: 1
	Interesse		= limpar_texto(Request("frmInteresse")) 	' Value: 1

	If id_idioma = 1 Then
		' Limpando o CPF
		CPF 			= limpar_texto(Request("frmCPF"))
		CPF				= Replace(CPF,".","")
		CPF				= Replace(CPF,"-","")
		If Len(CPF) <> 11 AND id_idioma = 1 Then
			Erro = Erro & "{ ""Erro"", ""02 - CPF Incorreto"" },"
		End If
	Else
		Passaporte		= limpar_texto(Request("frmCPF"))	' Numero do Passaporte
	End If

	response.write("ver CPF: " & CPF)
	'response.end()

	Nome 			= limpar_texto(Request("frmNome"))					' Value: MONICA
	NmCracha 		= Left(limpar_texto(Request("frmNmCracha")),27)		' Value: MONICA
	' Nova Implementacao - 04/05/2014 - Leandro Santiago
	' INICIO - Codigo do convite
	CodConvite 		= limpar_texto(Request("frmCodConvite"))			' Value: CD5087YYJJ
	If CodConvite = "-" OR Len(Trim(CodConvite)) = 0 Then CodConvite = "NULL"
	' FIM - Codigo do convite
	
	' Limpando o Data de Nascimento
	DtNasc			= limpar_texto(Request("frmDtNasc")) 		' Value: 28/01/1941
	DtNasc			= Replace(DtNasc,"/","") 					' Value: 28/01/1941
	If Len(DtNasc) <> 8 Then
		 Erro = Erro & "{ ""Erro"", ""04 - Data de Nascimento Incorreta"" },"
	End If
	
	Sexo			= limpar_texto(Request("frmSexo"))			' Value: 1
	Cargo			= limpar_texto(Request("frmCargo")) 		' Value: 11
	CargoOutros		= limpar_texto(Request("frmCargoOutros"))	' Value: ''
	'If CargoOutros = "-1" Then CargoOutros = "NULL"
	
	Depto 			= limpar_texto(Request("frmDepto"))			' Value: 1
	DeptoOutros		= limpar_texto(Request("frmDeptoOutros"))	' Value: ''
	SubCargo 		= limpar_texto(Request("frmSubCargo"))		' Value: 1
	If SubCargo = "-" OR Len(Trim(SubCargo)) = 0 Then SubCargo = "NULL"
	
	'response.write("SUBCARGO: " & SubCargo)
	
	DDI				= limpar_texto(Request("frmDDI")) 			' Value: 05
	DDD				= limpar_texto(Request("frmDDD")) 			' Value: 222
	
	' Limpando o Telefone 1
	Telefone		= limpar_texto(Request("frmTelefone")) 		' Value: 2875-8752
	Telefone		= Replace(Telefone,"-","")			 		' Value: 2875-8752
	If Len(Telefone) <> 8 AND id_idioma = 1 Then
		 Erro = Erro & "{ ""Erro"", ""05 - Telefone Incorreto"" },"
	End If
	TelefoneTipo 	= limpar_texto(Request("frmTipo"))			' Value: 1
	
	' Tratar SMS
	TelefoneSMS		= limpar_texto(Request("frmSMS"))			' Value: 1
	If Len(Trim(TelefoneSMS)) = 0 Then TelefoneSMS = 0
	
	
	Ramal			= limpar_texto(Request("frmRamal"))			' Value: 12345
	
	DDI2 			= limpar_texto(Request("frmDDI2"))			' Value: 11
	DDD2			= limpar_texto(Request("frmDDD2")) 			' Value: 111
	
	' Limpando o Telefone 2
	Telefone2		= limpar_texto(Request("frmTelefone2"))		' Value: 1111-1111
	Telefone2		= Replace(Telefone2,"-","")					' Value: 1111-1111
	If Len(Telefone2) <> 8 AND id_idioma = 1 Then
		 Erro = Erro & "{ ""Erro"", ""06 - Telefone 2 Incorreto"" },"
	End If
	TelefoneTipo2	= limpar_texto(Request("frmTipo2"))			' Value: 2
	
	' Tratar SMS2
	TelefoneSMS2	= limpar_texto(Request("frmSMS2"))			' Value: 1
	If Len(Trim(TelefoneSMS2)) = 0 Then TelefoneSMS2 = 0
	
	Ramal2			= limpar_texto(Request("frmRamal2"))		' Value: 12345
	
	DDIEmpresa		= limpar_texto(Request("frmDDIEmpresa"))			' Value: 11
	DDDEmpresa		= limpar_texto(Request("frmDDDEmpresa")) 			' Value: 111
	
	' Limpando o Telefone 2
	TelefoneEmpresa	= limpar_texto(Request("frmTelefoneEmpresa"))		' Value: 1111-1111
	TelefoneEmpresa	= Replace(TelefoneEmpresa,"-","")					' Value: 1111-1111
	If Len(TelefoneEmpresa) <> 8 AND id_idioma = 1 Then
		 Erro = Erro & "{ ""Erro"", ""06 - Telefone Empresa Incorreto"" },"
	End If
	TelefoneTipoEmpresa	= limpar_texto(Request("frmTipoEmpresa"))' Value: 2
	
	' Tratar SMS2
	TelefoneSMSEmpresa	= limpar_texto(Request("frmSMSEmpresa"))' Value: 1
	If Len(Trim(TelefoneSMSEmpresa)) = 0 Then TelefoneSMSEmpresa = 0
	
	RamalEmpresa			= limpar_texto(Request("frmRamalEmpresa"))		' Value: 12345
	
	Email 			= Lcase(limpar_texto(Request("frmEmail")))	' Value: monica.mendes@btsmedia.biz
	EmailConf 		= Lcase(limpar_texto(Request("frmEmailConf"))) ' Value: monica.mendes@btsmedia.biz
	If Email <> EmailConf Then
		 Erro = Erro & "{ ""Erro"", ""07 - E-mails não são iguais"" },"
	End If
	
	Newsletter		= limpar_texto(Request("frmNewsletter"))
	If Len(Trim(Newsletter)) = 0 Then
		Newsletter = 0
	End If	
	
	
	TotPerguntas 	= limpar_texto(Request("frmTotPerguntas"))

' Retorno das Validações do form
If Len(Erro) > 0 Then
	Retorno = "0"
Else
	Retorno = "1"
End If
%>