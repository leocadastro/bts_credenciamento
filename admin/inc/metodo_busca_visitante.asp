<%
	' Verificando Visitante
	SQL_Verificar_Visitante =	"Select " &_
								"	V.*, " &_
								"	C.Cargo_PTB as Cargo, " &_
								"	S.SubCargo_PTB as Subcargo, " &_
								"	D.Depto_PTB as Depto " &_
								"From Visitantes  as V " &_
								"Left Join Cargo as C  " &_
								"	ON C.ID_Cargo = V.ID_Cargo  " &_
								"Left Join SubCargo as S " &_
								"	On S.ID_Subcargo = V.ID_SubCargo " &_
								"Left Join Depto as D " &_
								"	ON D.ID_Depto = V.ID_Depto " &_
								"Where " &_ 
								"	ID_Visitante = " & id_visitante
	'response.write("<b>SQL_Verificar_Visitante</b><br>" & SQL_Verificar_Visitante & "<hr>")
	Set RS_Verificar_Visitante = Server.CreateObject("ADODB.Recordset")
	RS_Verificar_Visitante.Open SQL_Verificar_Visitante, Conexao

	If not RS_Verificar_Visitante.BOF or not RS_Verificar_Visitante.EOF Then
		CPF 					= RS_Verificar_Visitante("CPF")
	    CPFMask 				= Mid(CPF,1,3) & "." & Mid(CPF,4,3) & "." & Mid(CPF,7,3) & "-" & Mid(CPF,10,2)
		Passaporte				= RS_Verificar_Visitante("Passaporte")
		Nome_Completo 			= RS_Verificar_Visitante("Nome_Completo")
		Nome_Credencial 		= RS_Verificar_Visitante("Nome_Credencial")
		Data_Nasc 				= RS_Verificar_Visitante("Data_Nasc")
		DataMask				= Mid(Data_Nasc,1,2) & "/" & Mid(Data_Nasc,3,2) & "/" & Mid(Data_Nasc,5,4)
		Sexo 					= RS_Verificar_Visitante("Sexo")
		
		Select Case Sexo
			Case "1"
				Sexo = "Feminino"
			Case "0"
				Sexo = "Masculino"
		End Select
		
		Email 					= RS_Verificar_Visitante("Email")
		Newsletter 				= RS_Verificar_Visitante("Newsletter")
		ID_Cargo 				= RS_Verificar_Visitante("ID_Cargo")
		Cargo 					= RS_Verificar_Visitante("Cargo")		
		Cargo_Outros 			= RS_Verificar_Visitante("Cargo_Outros")
		SubCargo 				= RS_Verificar_Visitante("SubCargo")
		ID_SubCargo 			= RS_Verificar_Visitante("ID_SubCargo")
		SubCargo_Outros 		= RS_Verificar_Visitante("SubCargo_Outros")
		Depto 					= RS_Verificar_Visitante("Depto")
		ID_Depto 				= RS_Verificar_Visitante("ID_Depto")
		Depto_Outros			= RS_Verificar_Visitante("Depto_Outros")
		Data_Cadastro			= RS_Verificar_Visitante("Data_Cadastro")
	End If

	' Verificando Telefone
	SQL_Verificar_Telefone	=	"Select " &_
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
								"	AND ID_Visitante = " & id_visitante
	'response.write("<b>RS_Verificar_Telefone</b><br>" & SQL_Verificar_Telefone & "<hr>")
	Set RS_Verificar_Telefone = Server.CreateObject("ADODB.Recordset")
	RS_Verificar_Telefone.Open SQL_Verificar_Telefone, Conexao

	If not RS_Verificar_Telefone.BOF or not RS_Verificar_Telefone.EOF Then
		While not RS_Verificar_Telefone.EOF
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
			TelefoneVisitante	= DDIEmpresa & " (" & DDDEmpresa & ") " & NumeroEmpresa & " " & RamalEmpresa & " - " & TipoEmpresa
			RS_Verificar_Telefone.MoveNext
			If not RS_Verificar_Telefone.EOF Then TelefoneVisitante = TelefoneVisitante & "<br>"
		Wend
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
					            "	AND ID_Visitante = " & id_visitante
	'response.write("<b>RS_Verificar_Endereco</b><br>" & SQL_Verificar_Endereco & "<hr>")
	Set RS_Verificar_Endereco = Server.CreateObject("ADODB.Recordset")
	RS_Verificar_Endereco.Open SQL_Verificar_Endereco, Conexao

	If not RS_Verificar_Endereco.BOF or not RS_Verificar_Endereco.EOF Then
		CEP     = RS_Verificar_Endereco("CEP")
		CEPMask = Mid(CEP,1,5) & "-" & Mid(CEP,6,8)

		Endereco_Visitante = "<strong>" & CEPMask & "</strong><br/>"
		Endereco_Visitante = Endereco_Visitante & RS_Verificar_Endereco("Endereco") & ", " & RS_Verificar_Endereco("Numero") & " " & RS_Verificar_Endereco("Complemento") & "<br/>"
		Endereco_Visitante = Endereco_Visitante & RS_Verificar_Endereco("Bairro") & " - " & RS_Verificar_Endereco("Cidade") & "<br/>"
		Endereco_Visitante = Endereco_Visitante & RS_Verificar_Endereco("Sigla") & " - " & RS_Verificar_Endereco("Pais") & "<br/>"
	End If

%>