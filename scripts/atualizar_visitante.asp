<%
		'=======================================================================
		SQL_Verificar_Visitante	=	"Select " &_
									"	ID_Visitante " &_
									"From Visitantes " &_
									"Where " &_
									"	ID_Visitante = " & ID_Visitante 

		 response.write("<b>SQL_Verificar_Visitante</b><br>" & SQL_Verificar_Visitante & "<hr>")

		Set RS_Verificar_Visitante = Server.CreateObject("ADODB.Recordset")
		RS_Verificar_Visitante.Open SQL_Verificar_Visitante, Conexao
		'=======================================================================
		
		'=======================================================================
		'Se existe Atualizar
		If not RS_Verificar_Visitante.BOF or not RS_Verificar_Visitante.EOF Then
			RS_Verificar_Visitante.Close

			SQL_Atualizar_Visitante = 	"Update Visitantes " &_
										"Set " &_
										"	Nome_Completo			= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Nome,150) 			& "')) " &_
										"	,Nome_Credencial 		= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(NmCracha,27) 		& "')) " &_
										"	,Data_Nasc 				= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(DtNasc,8) 			& "')) " &_
										"	,Sexo			 		= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Sexo,1) 			& "')) " &_
										"	,Email 					= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Email,150) 		& "')) " &_
										"	,Newsletter 			= " & Newsletter & " " &_
										"	,ID_Cargo 				= " & Cargo & " " &_
										"	,Cargo_Outros 			= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(CargoOutros,50) 	& "')) " &_
										"	,ID_SubCargo 			= " & SubCargo & " " &_
										"	,SubCargo_Outros 		= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(SubCargoOutros,50)	& "')) " &_
										"	,ID_Depto 				= " & Depto & " " &_
										"	,Depto_Outros 			= Upper(dbo.sp_rm_accent_pt_latin1('" & Left(DeptoOutros,50)	& "')) " &_
										"	,Data_Atualizacao 		= getDate() " &_
										"	,Senha 					= " & Novo_ID_Rel_Cadastro & " " &_
										"Where ID_Visitante = " & ID_Visitante
									
			response.write("<b>SQL_Atualizar_Visitante</b><br>" & SQL_Atualizar_Visitante & "<hr>")
									
			Set RS_Atualizar_Visitante = Conexao.Execute(SQL_Atualizar_Visitante)
		End If
		'=======================================================================


		'=======================================================================
		' Verificar Telefone
		SQL_Telefone =	"Select ID_Relacionamento_Telefone " &_
						"From Relacionamento_Telefones " &_
						"Where ID_Visitante = " & ID_Visitante

		 response.write("<b>SQL_Telefone</b><br>" & SQL_Telefone & "<hr>")
								
		Set RS_Telefone = Server.CreateObject("ADODB.Recordset")
		RS_Telefone.Open SQL_Telefone, Conexao
		'=======================================================================
		
		'=======================================================================
		' Se existir Atualizar
		If not RS_Telefone.BOF or not RS_Telefone.EOF Then
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
										"	" & ID_Visitante & " " &_
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
			
			While not RS_Telefone.EOF
				ID_Tel = RS_Telefone("ID_Relacionamento_Telefone")
				SQL_Apagar_Telefone = 	"Delete From Relacionamento_Telefones Where ID_Relacionamento_Telefone = " & ID_tel
				' Executando 
				Conexao.Execute(SQL_Apagar_Telefone)
				
				RS_Telefone.MoveNext
			Wend
			RS_Telefone.Close
		'=======================================================================
		' Se não Cadastrar
		Else 
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
										"	" & ID_Visitante & " " &_
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
		End If
		'=======================================================================
		
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
										"	" & ID_Visitante & " " &_
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

		' Verificando se o Formulario e de PF para Atualizar o endereco
		If ID_Formulario = 4 then 

			'=======================================================================
			' Verificar Endereco
			SQL_Endereco =	"Select ID_Relacionamento_Endereco " &_
							"From Relacionamento_Enderecos " &_
							"Where ID_Visitante = " & ID_Visitante

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
											"	AND ID_Visitante = " & ID_Visitante
				 response.write("<b>SQL_Atualizar_Endereco</b><br>" & SQL_Atualizar_Endereco & "<hr>")
				' Executando Gravação
				Set RS_Atualizar_Endereco = Conexao.Execute(SQL_Atualizar_Endereco)
			Else
				'=======================================================================
				' Inserir Endereco da EMPRESA
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
										"	" & ID_Visitante & " " &_
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

		End If
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
								" 	ID_Visitante = " & ID_Visitante & " " &_
								"Where " &_
								"	ID_Relacionamento_Cadastro = " & Novo_ID_Rel_Cadastro
		 response.write("<b>SQL_Upd_Rel_Cadastro</b><br>" & SQL_Upd_Rel_Cadastro & "<hr>")
		' Executando Gravação
		Set RS_Upd_Rel_Cadastro = Conexao.Execute(SQL_Upd_Rel_Cadastro)
		'=======================================================================

response.write("<span style='background-color:#0F0'><b>Visitante ATUALIZADO com Sucesso</b></span><hr>")
%>