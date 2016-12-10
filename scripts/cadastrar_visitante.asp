<%
		'=======================================================================
		' Inserir VISITANTE
		SQL_Cad_Visitante = 	"SET NOCOUNT ON;" &_
								" " & vbCrLf & " " &_
								"INSERT INTO Visitantes " &_
								"	( " &_
								"	CPF " &_
								"	,Passaporte " &_
								"	,Nome_Completo " &_
								"	,Nome_Credencial " &_
								"	,Data_Nasc " &_
								"	,Sexo " &_
								"	,Email " &_
								"	,Newsletter " &_
								"	,ID_Cargo " &_
								"	,Cargo_Outros " &_
								"	,ID_SubCargo " &_
								"	,SubCargo_Outros " &_
								"	,ID_Depto " &_
								"	,Depto_Outros " &_
								"	,Senha " &_
								"	) " &_
								"VALUES " &_
								"	( " &_
								"	'" & Left(CPF,11) & "' " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Passaporte,50) &"')) " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Nome,150) &"')) " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(NmCracha,27) & "')) " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(DtNasc,8) & "')) " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Sexo,1) & "')) " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Email,150) & "')) " &_
								"	," & Newsletter & " " &_
								"	," & Cargo & " " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(CargoOutros,50) & "')) " &_
								"	," & SubCargo & " " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(SubCargoOutros,50) & "')) " &_
								"	," & Depto & " " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(DeptoOutros,50) & "')) " &_
								"	,'" & Novo_ID_Rel_Cadastro & "'" &_
								"	); " &_
								" " & vbCrLf & " " &_
								"SELECT @@Identity as NovoID; "

		 response.write("<b>SQL_Cad_Visitante</b><br>" & SQL_Cad_Visitante & "<hr>")

		' Executando Gravação com Retorno do ID
		Set RS_Cad_Visitante = Conexao.Execute(SQL_Cad_Visitante)
		Novo_ID_Visitante = RS_Cad_Visitante.Fields("NovoID").value
		Set RS_Cad_Visitante = Nothing
		response.write("Novo_ID_Visitante: " & Novo_ID_Visitante)
		'=======================================================================

		' Verificando se o Formulario e de PF para Atualizar o endereco
		If ID_Formulario = 4 then 

			'=======================================================================
			' Inserir Endereco da Visitantes
			SQL_Cad_End_Visitante = 	"INSERT INTO Relacionamento_Enderecos " &_
									"	( " &_
									"	ID_Visitante " &_
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
									"	" & Novo_ID_Visitante & " " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(CEP,12) & "')) " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Endereco,200) & "')) " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Numero,20) & "')) " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Complemento,50) & "')) " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Bairro, 200) & "')) " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Cidade, 200) & "')) " &_
									"	," & Estado & " " &_
									"	," & Pais & " " &_
									"	);"
			response.write("<b>SQL_Cad_End_Visitante</b><br>" & SQL_Cad_End_Visitante & "<hr>")
			' Executando Gravação
			Set RS_Cad_End_Visitante = Conexao.Execute(SQL_Cad_End_Visitante)
			'=======================================================================

		End If	

		'=======================================================================
		' Inserir TELEFONES DO VISITANTE
		SQL_Cad_Tel_Visitante = 	"INSERT INTO Relacionamento_Telefones " &_
									"	( " &_
									"	ID_Visitante " &_
									"	,ID_Tipo_Telefone " &_
									"	,DDI " &_
									"	,DDD " &_
									"	,Numero " &_
									"	,Ramal " &_
									"	,SMS " &_
									"	) " &_
									"VALUES " &_
									"	( " &_
									"	" & Novo_ID_Visitante & " " &_
									"	," & TelefoneTipo & " " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(DDI,5) & "')) " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(DDD,5) & "')) " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Telefone,15) & "')) " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Ramal,5) & "')) " &_
									"	," & TelefoneSMS & " " &_
									"	)"
		 response.write("<b>SQL_Cad_Tel_Visitante</b><br>" & SQL_Cad_Tel_Visitante & "<hr>")
		' Executando Gravação
		Set RS_Cad_Tel_Visitante = Conexao.Execute(SQL_Cad_Tel_Visitante)
		'=======================================================================

		If Len(Telefone2) > 0 Then

			'=======================================================================
			' Inserir TELEFONES DO VISITANTE
			SQL_Cad_Tel_Visitante = 	"INSERT INTO Relacionamento_Telefones " &_
										"	( " &_
										"	ID_Visitante " &_
										"	,ID_Tipo_Telefone " &_
										"	,DDI " &_
										"	,DDD " &_
										"	,Numero " &_
										"	,Ramal " &_
										"	,SMS " &_
										"	) " &_
										"VALUES " &_
										"	( " &_
										"	" & Novo_ID_Visitante & " " &_
										"	," & TelefoneTipo2 & " " &_
										"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(DDI2,5) & "')) " &_
										"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(DDD2,5) & "')) " &_
										"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Telefone2,15) & "')) " &_
										"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Ramal2,5) & "')) " &_
										"	," & TelefoneSMS2 & " " &_
										"	)"
			 response.write("<b>SQL_Cad_Tel_Visitante</b><br>" & SQL_Cad_Tel_Visitante & "<hr>")
			' Executando Gravação
			Set RS_Cad_Tel_Visitante = Conexao.Execute(SQL_Cad_Tel_Visitante)
			'=======================================================================

		End if

		'=======================================================================
		' Inserir PERGUNTAS
		' Existe total de perguntas ?
		If Len(TotPerguntas) > 0 Then
			' Loop na quantidade	
			For x = 1 To TotPerguntas
				ID_Pergunta = limpar_texto(Request("ID_Pergunta_" & x))
				Lista_Perguntas = Split(limpar_texto(Request("frmPergunta_" & x)),",")

				' Loop nos valores
				For y = Lbound(Lista_Perguntas) to Ubound(Lista_Perguntas)

					SQL_Cad_Perguntas = 	"INSERT INTO Relacionamento_Perguntas " &_
											"	( " &_
											"	ID_Relacionamento_Cadastro " &_
											"	,ID_Perguntas " &_
											"	,ID_Opcoes " &_
											"	,Texto " &_
											"	) " &_
											"VALUES " &_
											"	(" &_
											"	" & Novo_ID_Rel_Cadastro & ", " &_
											"	" & ID_Pergunta & ", " &_
											"	" & Lista_Perguntas(y) & ", " &_
											"	'');"
					 response.write("<b>SQL_Cad_Perguntas x(" & x & ") / y(" & y & ") / val(" & Lista_Perguntas(y) & ") </b><br>" & SQL_Cad_Perguntas & "<hr>")
					' Executando Gravação
					Set RS_Cad_Perguntas = Conexao.Execute(SQL_Cad_Perguntas)
				Next
			Next
		End If
		'=======================================================================

		'=======================================================================
		' Atualizar RELACIONAMENTO CADASTRO
		SQL_Upd_Rel_Cadastro =	"Update Relacionamento_Cadastro " &_
								"Set " &_
								" 	ID_Visitante = " & Novo_ID_Visitante & ", " &_
								"	CodConvite = '" & CodConvite & "' " &_
								"Where " &_
								"	ID_Relacionamento_Cadastro = " & Novo_ID_Rel_Cadastro
		 response.write("<b>SQL_Upd_Rel_Cadastro</b><br>" & SQL_Upd_Rel_Cadastro & "<hr>")
		' Executando Gravação
		Set RS_Upd_Rel_Cadastro = Conexao.Execute(SQL_Upd_Rel_Cadastro)
		'=======================================================================
response.write("<span style='background-color:#0F0'><b>Visitante Cadastrado com Sucesso</b></span><hr>")

%>