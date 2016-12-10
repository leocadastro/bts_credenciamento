<%
	' Verificando Empresa
	SQL_Verificar_Empresa =	"Select " &_
							"	* " &_
							"From Empresas " &_
							"Where " &_
							"	ID_Empresa = " & id_empresa 
	'response.write("<b>SQL_Verificar_Empresa</b><br>" & SQL_Verificar_Empresa & "<hr>")
	Set RS_Verificar_Empresa = Server.CreateObject("ADODB.Recordset")
	RS_Verificar_Empresa.Open SQL_Verificar_Empresa, Conexao

	If not RS_Verificar_Empresa.BOF or not RS_Verificar_Empresa.EOF Then
		CNPJ 				= RS_Verificar_Empresa("CNPJ")
	    CNPJMask 			= Mid(CNPJ,1,2) & "." & Mid(CNPJ,3,3) & "." & Mid(CNPJ,6,3) & "/" & Mid(CNPJ,9,4) & "-" & Mid(CNPJ,13,2)
		Funcionario_Qtde 	= RS_Verificar_Empresa("ID_Funcionarios_Qtde")
		Razao 				= RS_Verificar_Empresa("Razao_Social")
		Fantasia 			= RS_Verificar_Empresa("Nome_Fantasia")
		Produto 			= RS_Verificar_Empresa("Principal_Produto")
		Site 				= RS_Verificar_Empresa("site")
		Email 				= RS_Verificar_Empresa("Email")
		Newsletter 			= RS_Verificar_Empresa("Newsletter")
		Presidente 			= RS_Verificar_Empresa("Presidente")
		Reitor 				= RS_Verificar_Empresa("Reitor")
		Senha 				= RS_Verificar_Empresa("Senha")
		Data_Cadastro 		= RS_Verificar_Empresa("Data_Cadastro")
		Data_Atualizacao 	= RS_Verificar_Empresa("Data_Atualizacao")
		RS_Verificar_Empresa.Close
	End If

	' Verificando o Ramo
	SQL_Verificar_Ramo =  	"Select " &_
				            "   Ramo_PTB as Ramo " &_
				            "   ,Ramo_Outros as Outros " &_
				            "From " &_
				            "   RamodeAtividade as RA " &_
				            "Inner Join " &_
				            "   Relacionamento_Ramo as RR " &_
				            "   ON RA.ID_Ramo = RR.ID_Ramo " &_
				            "Where RR.ID_Empresa = " & id_empresa
	'response.write("<b>SQL_Verificar_Ramo</b><br>" & SQL_Verificar_Ramo & "<hr>")
	Set RS_Verificar_Ramo = Server.CreateObject("ADODB.Recordset")
	RS_Verificar_Ramo.Open SQL_Verificar_Ramo, Conexao

	If not RS_Verificar_Ramo.BOF or not RS_Verificar_Ramo.EOF Then
		If isNull(RS_Verificar_Ramo("Outros")) = false Then
			Ramo          = "Outros: " & RS_Verificar_Ramo("Outros")
			SemAtividade  = True
		Else
			Ramo = RS_Verificar_Ramo("Ramo")
		End If
		RS_Verificar_Ramo.Close
	End If

	' Verificando a Atividade
	SQL_Verificar_Atividade = 	"Select " &_
			                    "  AE.Atividade_PTB as Atividade  " &_
			                    "  ,RA.Atividade_Outros as Outros  " &_
			                    "From  " &_
			                    "  AtividadeEconomica as AE " &_
			                    "Inner Join  " &_
			                    "  Relacionamento_Atividade as RA " &_
			                    "  ON AE.ID_Atividade = RA.ID_Atividade " &_
			                    "Where  " &_
			                    "  RA.ID_Empresa = " & id_empresa
	'response.write("<b>SQL_Verificar_Atividade</b><br>" & SQL_Verificar_Atividade & "<hr>")
	Set RS_Verificar_Atividade = Server.CreateObject("ADODB.Recordset")
	RS_Verificar_Atividade.Open SQL_Verificar_Atividade, Conexao

	' Verificando o retorno da QUERY
	If not RS_Verificar_Atividade.BOF or not RS_Verificar_Atividade.EOF Then
		If isNull(RS_Verificar_Atividade("Outros")) = false Then
			Atividade = "Outros: " & RS_Verificar_Atividade("Outros")        
		Elseif SemAtividade = True Then
			Atividade = ""
		Else
			Atividade = RS_Verificar_Atividade("Atividade")
		End If
		RS_Verificar_Atividade.Close
	End If

	' Verificando Telefone
	SQL_Verificar_Telefone	=	"Select Top 1" &_
								"	DDI " &_
								"	,DDD " &_
								"	,Numero " &_
								"	,Ramal " &_
								"	,SMS " &_
								"	,RT.ID_Tipo_Telefone " &_
	      						"	,Tipo_PTB " &_
								"From " &_
								"	Relacionamento_Telefones as RT " &_
								"INNER JOIN " &_
								"	Tipo_Telefone as TT " &_
								"	ON TT.ID_Tipo_Telefone = RT.ID_Tipo_Telefone " &_
								"Where " &_
								"	RT.Ativo = 1 " &_	
								"	AND ID_Empresa = " & id_empresa
	'response.write("<b>RS_Verificar_Telefone</b><br>" & SQL_Verificar_Telefone & "<hr>")
	Set RS_Verificar_Telefone = Server.CreateObject("ADODB.Recordset")
	RS_Verificar_Telefone.Open SQL_Verificar_Telefone, Conexao

	If not RS_Verificar_Telefone.BOF or not RS_Verificar_Telefone.EOF Then
		DDIEmpresa             = RS_Verificar_Telefone("DDI")
		DDDEmpresa             = RS_Verificar_Telefone("DDD")
		NumeroEmpresa          = RS_Verificar_Telefone("Numero")
		RamalEmpresa           = RS_Verificar_Telefone("Ramal")
		SMSEmpresa             = RS_Verificar_Telefone("SMS")
		TipoEmpresa            = RS_Verificar_Telefone("Tipo_PTB")
		VerTipoTelefoneEmpresa = RS_Verificar_Telefone("ID_Tipo_Telefone")

		' Verificando o tipo de telefone
		If VerTipoTelefoneEmpresa = "3" Then
			TituloSMS = "SMS"
			' Verificando se o tipo de telefone e Celular
			If Len(SMSEmpresa) > 0 Then
				RecebeSMSEmpresa = "Sim"
			Else
				RecebeSMSEmpresa = "Não"
			End If
		End If
		TelefoneEmpresa	= DDIEmpresa & " (" & DDDEmpresa & ") " & NumeroEmpresa & " " & RamalEmpresa & " - " & TipoEmpresa
		RS_Verificar_Telefone.Close
	End If
			
	' Verificando Endereco
	SQL_Verificar_Endereco =	"Select " &_
					            "	CEP " &_
					            "	,Endereco " &_
					            "	,Numero " &_
					            "	,Complemento " &_
					            "	,Bairro " &_
					            "	,Cidade " &_
					            "	,RE.ID_UF " &_
					            "	,Sigla " &_
					            "	,RE.ID_Pais " &_
					            "	,Pais_PTB as Pais " &_
					            "From " &_
					            "	Relacionamento_Enderecos as RE " &_
					            "INNER JOIN " &_
					            "	UF as UF " &_
					            "	ON UF.ID_UF = RE.ID_UF " &_
					            "INNER JOIN " &_
					            "	Pais as PA " &_
					            "	ON PA.ID_Pais = RE.ID_Pais " &_
					            "Where " &_
					            "	RE.Ativo = 1 " &_ 
					            "	AND ID_Empresa = " & id_empresa
	'response.write("<b>RS_Verificar_Endereco</b><br>" & SQL_Verificar_Endereco & "<hr>")
	Set RS_Verificar_Endereco = Server.CreateObject("ADODB.Recordset")
	RS_Verificar_Endereco.Open SQL_Verificar_Endereco, Conexao

	If not RS_Verificar_Endereco.BOF or not RS_Verificar_Endereco.EOF Then
		CEP     = RS_Verificar_Endereco("CEP")
		CEPMask = Mid(CEP,1,5) & "-" & Mid(CEP,6,8)

		Endereco = "<strong>" & CEPMask & "</strong><br/>"
		Endereco = Endereco & RS_Verificar_Endereco("Endereco") & ", " & RS_Verificar_Endereco("Numero") & " " & RS_Verificar_Endereco("Complemento") & "<br/>"
		Endereco = Endereco & RS_Verificar_Endereco("Bairro") & " - " & RS_Verificar_Endereco("Cidade") & "<br/>"
		Endereco = Endereco & RS_Verificar_Endereco("Sigla") & " - " & RS_Verificar_Endereco("Pais") & "<br/>"
		RS_Verificar_Endereco.Close
	End If

	' Verificando Funcionarios
	if(len(Funcionario_Qtde) > 0) then
		SQL_Verificar_Funcionarios =	"Select " &_
										"	Funcionarios_Qtde_PTB " &_
										"From " &_
										"	Funcionarios_Qtde " &_
										"Where " &_
										"	Ativo = 1 " &_ 
										"	AND ID_Funcionarios_Qtde = " & Funcionario_Qtde
	else
		SQL_Verificar_Funcionarios =	"Select " &_
										"	Funcionarios_Qtde_PTB " &_
										"From " &_
										"	Funcionarios_Qtde " &_
										"Where " &_
										"	Ativo = 1 "
	end if
	'response.write("<b>SQL_Verificar_Funcionarios</b><br>" & SQL_Verificar_Funcionarios & "<hr>")
	Set RS_Verificar_Funcionarios = Server.CreateObject("ADODB.Recordset")
	RS_Verificar_Funcionarios.Open SQL_Verificar_Funcionarios, Conexao

	If not RS_Verificar_Funcionarios.BOF or not RS_Verificar_Funcionarios.EOF Then
		Funcionarios_Qtde = RS_Verificar_Funcionarios("Funcionarios_Qtde_PTB")
		RS_Verificar_Funcionarios.Close
	End If
%>