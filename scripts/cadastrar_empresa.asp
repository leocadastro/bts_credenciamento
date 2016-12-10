<%
		' #################################################################################
		' Inserir EMPRESA
		'	Inserir Endereço empresa
		'	Inserir Interesses na feira
		' Inserir VISITANTE
		'	Inserir Telefones 1 e 2
		' Inserir Relacionamento Feira > Empresa > Visitante
		' Inserir Perguntas com o ID do Relacionamento
		' #################################################################################
		
		'=======================================================================
		' Inserir EMPRESA
		'	ID_Formulario - Nome
		'	1 - Empresa
		'	2 - Entidades
		'	3 - Imprensa
		'	4 - Pessoa Física
		'	5 - Universidades
		'	6 - Alunos
		
		Select Case ID_Formulario
			'=======================================================================
			Case 1
				SQL_Cad_Empresa = 	"SET NOCOUNT ON;" &_
									" " & vbCrLf & " " &_
									"INSERT INTO Empresas " &_
									"	(ID_Formulario " &_
									"	,ID_Funcionarios_Qtde " &_
									"	,CNPJ " &_
									"	,Razao_Social " &_
									"	,Nome_Fantasia " &_
									"	,Site) " &_
									"VALUES " &_
									"	(" & ID_Formulario & ", " &_
									"	" & QtdFunc & ", " &_
									"	Upper(dbo.sp_rm_accent_pt_latin1('" & Left(CNPJ, 14) & "')), " &_
									"	Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Razao, 150) & "')), " &_
									"	Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Fantasia, 150) & "')), " &_
									"	Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Site, 100) & "')) " &_
									"	); " &_
									" " & vbCrLf & " " &_
									"SELECT @@Identity as NovoID; "
			'=======================================================================
			Case 2
				SQL_Cad_Empresa = 	"SET NOCOUNT ON;" &_
									" " & vbCrLf & " " &_
									"INSERT INTO Empresas " &_
									"	(ID_Formulario " &_
									"	,CNPJ " &_
									"	,Razao_Social " &_
									"	,Nome_Fantasia " &_
									"	,Presidente " &_
									"	,Site) " &_
									"VALUES " &_
									"	(" & ID_Formulario & ", " &_
									"	Upper(dbo.sp_rm_accent_pt_latin1('" & Left(CNPJ, 14) & "')), " &_
									"	Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Razao, 150) & "')), " &_
									"	Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Sigla, 150) & "')), " &_
									"	Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Resp, 100) & "')), " &_
									"	Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Site, 100) & "')) " &_
									"	); " &_
									" " & vbCrLf & " " &_
									"SELECT @@Identity as NovoID; "
			'=======================================================================
			Case 5
				SQL_Cad_Empresa = 	"SET NOCOUNT ON;" &_
									" " & vbCrLf & " " &_
									"INSERT INTO Empresas " &_
									"	(ID_Formulario " &_
									"	,CNPJ " &_
									"	,Razao_Social " &_
									"	,Nome_Fantasia " &_
									"	,Site " &_
									"	,Senha " &_
									"	,Reitor) " &_
									"VALUES " &_
									"	(" & ID_Formulario & ", " &_
									"	Upper(dbo.sp_rm_accent_pt_latin1('" & Left(CNPJ, 14) & "')), " &_
									"	Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Razao, 150) & "')), " &_
									"	Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Fantasia, 150) & "')), " &_
									"	Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Site, 100) & "')), " &_
									"	Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Senha, 20) & "')), " &_
									"	Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Resp, 100) & "')) " &_
									"	); " &_
									" " & vbCrLf & " " &_
									"SELECT @@Identity as NovoID; "
			'=======================================================================
		End Select

		 response.write("<b>SQL_Cad_Empresa</b><br>" & SQL_Cad_Empresa & "<hr>")

		'Executando Gravação com Retorno do ID
		Set RS_Cad_Empresa = Conexao.Execute(SQL_Cad_Empresa)
		Novo_ID_Empresa = RS_Cad_Empresa.Fields("NovoID").value
		Set RS_Cad_Empresa = Nothing
		response.write("Novo_ID_Empresa: " & Novo_ID_Empresa & "<br>")
		'=======================================================================

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
							"	 " & Novo_ID_Empresa & ", " &_
							"	 '" & Left(CodConvite, 15) & "' );" &_
							" " & vbCrLf & " " &_
							"SELECT @@Identity as NovoID; "

		 response.write("<b>SQL_Rel_Cadastro</b><br>" & SQL_Rel_Cadastro & "<hr>")

		' Executando Gravação com Retorno do ID
		Set RS_Rel_Cadastro = Conexao.Execute(SQL_Rel_Cadastro)
		Novo_ID_Rel_Cadastro = RS_Rel_Cadastro.Fields("NovoID").value
		Set RS_Rel_Cadastro = Nothing
		response.write("Novo_ID_Rel_Cadastro: " & Novo_ID_Rel_Cadastro & "<br>")
		'=======================================================================

		'=======================================================================
		' Verificar Produto, se não existir GRAVAR
		SQL_Verificar_Produto =	"Select ID "&_
								"From Relacionamento_Produtos " &_
								"Where " &_
								"	ID_Empresa = " & Novo_ID_Empresa & " " &_
								"	AND Principal_Produto = Upper(dbo.sp_rm_accent_pt_latin1('" & PriProdut & "'))"
		 response.write("<b>SQL_Verificar_Produto</b><br>" & SQL_Verificar_Produto & "<hr>")
		Set RS_Verificar_Produto = Server.CreateObject("ADODB.Recordset")
		RS_Verificar_Produto.Open SQL_Verificar_Produto, Conexao
		'=======================================================================

		'=======================================================================
		'Se NAO existe GRAVAR
		If RS_Verificar_Produto.BOF or RS_Verificar_Produto.EOF Then
			SQL_Cad_Produto = 	"Insert Into Relacionamento_Produtos " &_
								"(ID_Empresa, Principal_Produto) " &_
								"Values " &_
								"	( " &_
								"	" & Novo_ID_Empresa & " " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & PriProdut & "')) " &_
								"	)"
			' Executando Gravação
			Set SQL_Cad_Produto = Conexao.Execute(SQL_Cad_Produto)
		Else
			RS_Verificar_Produto.Close
		End IF
		'=======================================================================

		'=======================================================================
		' Inserir Os Interesses Selecionados
		response.write("Lista_Interesses: " & Interesse)
		Lista_Interesses = Split(Interesse,",")
		For i = Lbound(Lista_Interesses) to Ubound(Lista_Interesses)
			response.write("i: " & i & " - v: " & Lista_Interesses(i) & "<br>")
			SQL_Cad_Interesse = 	"INSERT INTO Relacionamento_InteresseFeira " &_
									"	(ID_Relacionamento_Cadastro " &_
									"	,ID_InteresseFeira) " &_
									"VALUES " &_
									"	(" & Novo_ID_Rel_Cadastro & ", " &_
									"	" & Lista_Interesses(i) & "); "

			 response.write("<b>SQL_Cad_Interesse</b><br>" & SQL_Cad_Interesse & "<hr>")
			' Executando Gravação
			Set RS_Cad_Interesse = Conexao.Execute(SQL_Cad_Interesse)
		Next
		'=======================================================================

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
								"	" & Novo_ID_Empresa & " " &_
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
		
		'=======================================================================
		' Inserir TELEFONES DO EMPRESA
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
									"	" & Novo_ID_Empresa & " " &_
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
		'=======================================================================
response.write("<span style='background-color:#0F0'><b>Empresa Cadastrada com Sucesso</b></span><hr>")
%>