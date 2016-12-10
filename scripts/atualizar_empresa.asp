<%
		' #################################################################################
		' Atualizar EMPRESA
		'	Atualizar Endereço empresa
		'	Atualizar Interesses na feira
		' Atualizar VISITANTE
		'	Atualizar Telefones 1 e 2
		' Atualizar Relacionamento Feira > Empresa > Visitante
		' Atualizar Perguntas com o ID do Relacionamento
		' #################################################################################

		'=======================================================================
		' Inserir Relacionamento Cadastro
		SQL_Rel_Cadastro = 	"SET NOCOUNT ON;" &_
							" " & vbCrLf & " " &_
							"INSERT INTO Relacionamento_Cadastro " &_
							"	(ID_Idioma " &_
							"	,ID_Edicao " &_
							"	,ID_Tipo_Credenciamento " &_
							"	,ID_Empresa " &_
							"	,CodConvite) " &_
							"VALUES " &_
							"	(" & id_idioma & ", " &_
							"	 " & id_edicao & ", " &_
							"	 " & id_tipo & ", " &_
							"	 " & id_empresa & ", " &_
							"	 '" & Left(CodConvite, 15) & "' );" &_
							" " & vbCrLf & " " &_
							"SELECT @@Identity as NovoID; "

		 response.write("<b>SQL_Rel_Cadastro</b><br>" & SQL_Rel_Cadastro & "<hr>")

		' Executando Gravação com Retorno do ID
		Set RS_Rel_Cadastro = Conexao.Execute(SQL_Rel_Cadastro)
		Novo_ID_Rel_Cadastro = RS_Rel_Cadastro.Fields("NovoID").value
		Set RS_Rel_Cadastro = Nothing
		response.write("Novo_ID_Rel_Cadastro: " & Novo_ID_Rel_Cadastro)
		'=======================================================================

		'=======================================================================
		SQL_Verificar_Empresa	=	"Select " &_
									"	ID_Empresa " &_
									"From Empresas " &_
									"Where " &_
									"	ID_Empresa = " & id_empresa 

		 response.write("<b>SQL_Verificar_Empresa</b><br>" & SQL_Verificar_Empresa & "<hr>")

		Set RS_Verificar_Empresa = Server.CreateObject("ADODB.Recordset")
		RS_Verificar_Empresa.Open SQL_Verificar_Empresa, Conexao
		'=======================================================================

		'=======================================================================
		'Se existe Atualizar
		If not RS_Verificar_Empresa.BOF or not RS_Verificar_Empresa.EOF Then
			'	ID_Formulario - Nome
			'	1 - Empresa
			'	2 - Entidades
			'	3 - Imprensa
			'	4 - Pessoa Física
			'	5 - Universidades
			'	6 - Alunos
			response.write("ID_Formulario: "  & ID_Formulario & "<hr>")
			Select Case ID_Formulario
				'=======================================================================
				Case 1
					SQL_Atualizar_Empresa = "Update Empresas " &_
											"Set " &_
											"	ID_Formulario			= " & ID_Formulario & ", " &_
											"	ID_Funcionarios_Qtde	= " & QtdFunc 		& ", " &_
											"	CNPJ 					= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(CNPJ, 14) 			& "')), " &_
											"	Razao_Social 			= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Razao, 150) 		& "')), " &_
											"	Nome_Fantasia 			= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Fantasia, 150) 	& "')), " &_
											"	Principal_Produto 		= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(PriProdut, 150) 	& "')), " &_
											"	Site 					= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Site, 100) 		& "')), " &_
											"	Data_Atualizacao 		= getDate() " &_
											"Where ID_Empresa = " & id_empresa
				'=======================================================================
				Case 2
					SQL_Atualizar_Empresa = "Update Empresas " &_
											"Set " &_
											"	ID_Formulario			= " & ID_Formulario & ", " &_
											"	CNPJ 					= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(CNPJ, 14) 			& "')), " &_
											"	Razao_Social 			= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Razao, 150) 		& "')), " &_
											"	Nome_Fantasia 			= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Sigla, 150) 		& "')), " &_
											"	Presidente				= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Resp, 100) 		& "')), " &_
											"	Site 					= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Site, 100) 		& "')), " &_
											"	Data_Atualizacao 		= getDate() " &_
											"Where ID_Empresa = " & id_empresa
				'=======================================================================
				Case 5
					SQL_Atualizar_Empresa = "Update Empresas " &_
											"Set " &_
											"	ID_Formulario			= " & ID_Formulario & ", " &_
											"	CNPJ 					= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(CNPJ, 14) 			& "')), " &_
											"	Razao_Social 			= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Razao, 150) 		& "')), " &_
											"	Nome_Fantasia 			= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Fantasia, 150) 		& "')), " &_
											"	Site 					= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Site, 100) 		& "')), " &_
											"	Senha 					= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Senha, 20) 		& "')), " &_
											"	Reitor					= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Resp, 100)	 		& "')), " &_
											"	Data_Atualizacao 		= getDate() " &_
											"Where ID_Empresa = " & id_empresa
				'=======================================================================
			End Select
									
			 response.write("<b>SQL_Atualizar_Empresa</b><br>" & SQL_Atualizar_Empresa & "<hr>")
									
			Set RS_Atualizar_Empresa = Conexao.Execute(SQL_Atualizar_Empresa)	
		Else
			
		End If
		'=======================================================================

		'=======================================================================
		' Verificar Telefone
		SQL_Telefone =	"Select ID_Relacionamento_Telefone " &_
						"From Relacionamento_Telefones " &_
						"Where ID_Empresa = " & id_empresa

		 response.write("<b>SQL_Telefone</b><br>" & SQL_Telefone & "<hr>")
								
		Set RS_Telefone = Server.CreateObject("ADODB.Recordset")
		RS_Telefone.Open SQL_Telefone, Conexao
		'=======================================================================
		
		'=======================================================================
		' Se existir Atualizar
		If not RS_Telefone.BOF or not RS_Telefone.EOF Then
			SQL_Atualizar_Telefone = 	"Update Relacionamento_Telefones " &_
										"Set " &_
										"	ID_Tipo_Telefone 	= " & TelefoneTipoEmpresa & " " &_
										"	,DDI 				= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(DDIEmpresa,5) & "')) " &_
										"	,DDD 				= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(DDDEmpresa,5) & "')) " &_
										"	,Numero				= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(TelefoneEmpresa,15) & "')) " &_
										"	,Ramal 				= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(RamalEmpresa,5) & "')) " &_
										"	,SMS 				= " & TelefoneSMSEmpresa & " " &_
										"	,Data_Atualizacao 	= getDate() " &_
										"Where ID_Empresa = " & id_empresa
			 response.write("<b>SQL_Atualizar_Telefone</b><br>" & SQL_Atualizar_Telefone & "<hr>")
			' Executando Gravação
			Set RS_Atualizar_Telefone = Conexao.Execute(SQL_Atualizar_Telefone)										
		'=======================================================================
		' Se não Cadastrar
		Else 
			SQL_Cad_Tel_Empresa = 	"INSERT INTO Relacionamento_Telefones " &_
									"	( " &_
									"	ID_Empresa " &_
									"	,ID_Tipo_Telefone " &_
									"	,DDI " &_
									"	,DDD " &_
									"	,Numero " &_
									"	,Ramal " &_
									"	,SMS " &_
									"	) " &_
									"VALUES " &_
									"	( " &_
									"	" & id_empresa & " " &_
									"	," & TelefoneTipoEmpresa & " " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(DDIEmpresa,5) & "')) " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(DDDEmpresa,5) & "')) " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(TelefoneEmpresa,15) & "')) " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(RamalEmpresa,5) & "')) " &_
									"	," & TelefoneSMSEmpresa & " " &_
									"	)"
			 response.write("<b>SQL_Cad_Tel_Empresa</b><br>" & SQL_Cad_Tel_Empresa & "<hr>")
			' Executando Gravação
			Set RS_Cad_Tel_Empresa = Conexao.Execute(SQL_Cad_Tel_Empresa)		
		End If
		'=======================================================================

		'=======================================================================
		' Verificar Endereco
		SQL_Endereco =	"Select ID_Relacionamento_Endereco " &_
						"From Relacionamento_Enderecos " &_
						"Where ID_Empresa = " & id_empresa

		 response.write("<b>SQL_Endereco</b><br>" & SQL_Endereco & "<hr>")
								
		Set RS_Endereco = Server.CreateObject("ADODB.Recordset")
		RS_Endereco.Open SQL_Endereco, Conexao
		'=======================================================================

		'=======================================================================
		If not RS_Endereco.BOF or not RS_Endereco.EOF Then
			'=======================================================================
			' Se existir Atualizar
			SQL_Atualizar_Endereco = 	"Update Relacionamento_Enderecos " &_
										"Set " &_
										"	CEP 				= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(CEP,12) & "')) " &_
										"	,Endereco 			= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Endereco,200) & "')) " &_
										"	,Numero 			= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Numero,20) & "')) " &_
										"	,Complemento 		= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Complemento,50) & "')) " &_
										"	,Bairro 			= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Bairro, 200) & "')) " &_
										"	,Cidade 			= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Cidade, 200) & "')) " &_
										"	,ID_UF 				= " & Estado & " " &_
										"	,ID_Pais 			= " & Pais & " " &_
										"	,UF_Exterior 		= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(UF_Exterior, 100) & "')) " &_
										"	,Data_Atualizacao 	= getDate() " &_
										"Where " &_
										"	ID_Relacionamento_Endereco = " & RS_Endereco("ID_Relacionamento_Endereco") & " " &_
										"	AND ID_Empresa = " & id_empresa
			 response.write("<b>SQL_Atualizar_Endereco</b><br>" & SQL_Atualizar_Endereco & "<hr>")
			' Executando Gravação
			Set RS_Atualizar_Endereco = Conexao.Execute(SQL_Atualizar_Endereco)
		Else
			'=======================================================================
			' Inserir Endereco da EMPRESA
			SQL_Cad_End_Empresa = 	"INSERT INTO Relacionamento_Enderecos " &_
									"	( " &_
									"	ID_Empresa " &_
									"	,CEP " &_
									"	,Endereco " &_
									"	,Numero " &_
									"	,Complemento " &_
									"	,Bairro " &_
									"	,Cidade " &_
									"	,ID_UF " &_
									"	,ID_Pais " &_
									"	) " &_
									"VALUES " &_
									"	( " &_
									"	" & id_empresa & " " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(CEP,12) & "')) " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Endereco,200) & "')) " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Numero,20) & "')) " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Complemento,50) & "')) " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Bairro, 200) & "')) " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Cidade, 200) & "')) " &_
									"	," & Estado & " " &_
									"	," & Pais & " " &_
									"	);"
			 response.write("<b>SQL_Cad_End_Empresa</b><br>" & SQL_Cad_End_Empresa & "<hr>")
			' Executando Gravação
			Set RS_Cad_End_Empresa = Conexao.Execute(SQL_Cad_End_Empresa)
			'=======================================================================
		End If
		'=======================================================================
		
		'=======================================================================
		' Inserir Ramos Selecionados
		Lista_Ramos = Split(OptRamo,",")
		'=======================================================================
		
		'=======================================================================
		' Inserir Os Interesses Selecionados
		response.write("Lista_Interesses: " & Interesse)
		Lista_Interesses = Split(Interesse,",")
		For i = Lbound(Lista_Interesses) to Ubound(Lista_Interesses)
			response.write("i: " & i & " - v: " & Lista_Interesses(i) & "<br>")
			'=======================================================================
			SQL_Verificar_Interesse =	"Select ID_Relacionamento_Cadastro " &_
										"From Relacionamento_InteresseFeira " &_ 
										"Where " &_
										"	ID_Relacionamento_Cadastro = " & Novo_ID_Rel_Cadastro & " " &_
										"	AND ID_InteresseFeira = " & Lista_Interesses(i)
			Set RS_Verificar_Interesse = Server.CreateObject("ADODB.Recordset")
			RS_Verificar_Interesse.Open SQL_Verificar_Interesse, Conexao
			'=======================================================================

			'=======================================================================
			' Se o RAMO ainda não foi cadastrado, GRAVAR
			If RS_Verificar_Interesse.BOF or RS_Verificar_Interesse.EOF Then
				SQL_Cad_Interesse = 	"INSERT INTO Relacionamento_InteresseFeira " &_
										"	(ID_Relacionamento_Cadastro " &_
										"	,ID_InteresseFeira) " &_
										"VALUES " &_
										"	(" & Novo_ID_Rel_Cadastro & ", " &_
										"	" & Lista_Interesses(i) & "); "

				 response.write("<b>SQL_Cad_Interesse</b><br>" & SQL_Cad_Interesse & "<hr>")
				' Executando Gravação
				Set RS_Cad_Interesse = Conexao.Execute(SQL_Cad_Interesse)
			Else
				RS_Verificar_Interesse.Close
			End If
		Next
		'=======================================================================
response.write("<span style='background-color:#0F0'><b>Empresa Atualizada com Sucesso</b></span><hr>")
%>