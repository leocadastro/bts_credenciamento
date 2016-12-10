<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/admin/inc/gravar_limpar_texto.asp"-->
<!--#include virtual="/scripts/ConsultarWebService.asp"-->
<%
Form_Busca 	= Limpar_Texto(Trim(Request("busca")))
Pedido		= Limpar_Texto(Request("pedido"))

'Response.Write Form_Busca
'Response.End

If Len(Trim(Form_Busca)) > 0 Then
'Response.Write Form_Busca
'Response.End
	'==================================================
	Set Conexao = Server.CreateObject("ADODB.Connection")
	Conexao.Open Application("cnn")
	'==================================================
	

	'SQL_Busca = "Select " &_
	'			"	RC.ID_Relacionamento_Cadastro as IRC " &_
	'			"	,RC.ID_Edicao " &_
	'			"	,RC.ID_Empresa " &_
	'			"	,RC.ID_Visitante " &_
	'			"	,V.Nome_Completo " &_
	'			"	,V.CPF " &_
	'			"	,V.Senha " &_
	'			"	,V.Data_Nasc " &_
	'			"	,V.Passaporte " &_
	'			"From Relacionamento_Cadastro as RC " &_
	'			"Inner Join Tipo_Credenciamento as TC ON TC.ID_Tipo_Credenciamento = RC.ID_Tipo_Credenciamento " &_
	'			"Inner Join Visitantes as V ON V.ID_Visitante = RC.ID_Visitante " &_
	'			"Where  " &_
	'			"	RC.ID_Tipo_Credenciamento in (10,11,12)	/* Pessoa Fisica PTB, ESP e ENG	*/ " &_
	'			"	AND TC.ID_Idioma =  1 		/* Idioma	*/ " &_
	'			"	AND (V.CPF = '" & Form_Busca & "' OR V.Passaporte = '" & Form_Busca & "') " &_
	'			"	AND RC.ID_Edicao = " & Session("cliente_edicao")


	'==========Alterado Luiz Ricardo
	SQL_Busca = "Select " &_
				"	RC.ID_Relacionamento_Cadastro as IRC " &_
				"	,RC.ID_Edicao " &_
				"	,RC.ID_Empresa " &_
				"	,RC.ID_Visitante " &_
				"	,V.Nome_Completo " &_
				"	,V.CPF " &_
				"	,V.Senha " &_
				"	,V.Data_Nasc " &_
				"	,V.Passaporte " &_
				"From Relacionamento_Cadastro as RC " &_
				"Inner Join Tipo_Credenciamento as TC ON TC.ID_Tipo_Credenciamento = RC.ID_Tipo_Credenciamento " &_
				"Inner Join Visitantes as V ON V.ID_Visitante = RC.ID_Visitante " &_
				"Where  " &_
				"	TC.ID_Idioma =  1 		/* Idioma	*/ " &_
				"	AND (V.CPF = '" & Form_Busca & "' OR V.Email = '" & Form_Busca & "') " &_
				"	AND RC.ID_Edicao = " & Session("cliente_edicao")
	Set RS_Busca = Server.CreateObject("ADODB.Recordset")
	RS_Busca.Open SQL_Busca, Conexao, 3, 3
	'Response.Write(SQL_Busca & "<br>")
			'Response.end
	'// Se o visitante não foi localizado
	If Rs_Busca.EOF Then
		'// Importar dados do Ws
		Id_Visitante = ConsultarWS(Form_Busca)
		'response.write Id_Visitante
		'response.end
		'// Processar novamente o RecordSet
		Set RS_Busca = Conexao.Execute(SQL_Busca)
	End If

		If Not Rs_Busca.Eof Then
				
			'cortesia = 	ConsultarCortesia(Form_Busca)
			cortesia = "false"
			'response.write cortesia = "false"  
			'response.end
			if cortesia = "false" then
			' Luiz Ricardo 13/03 - Buscar por CPF
			' Busca apenas em meu carrinho para que eu não possa inserir novamente	
			SQL_Busca_Carrinho = 	"Select " &_
									"	C.ID_Carrinho,  " &_
									"	C.ID_Visitante,  " &_
									"	C.ID_Pedido,  " &_
									"	C.ID_Rel_Cadastro, " &_
									"	P.Status_Pedido, " &_
									"	V.Nome_Completo, " &_
									"	V.CPF, " &_
									"	V.Passaporte, " &_
									"	P.ID_Edicao " &_
									"From  Pedidos_Carrinho  As C " &_
									"Inner Join Visitantes As V On V.ID_Visitante = C.ID_Visitante " &_
									"Inner Join Pedidos As P On P.ID_Pedido = C.ID_Pedido " &_
									"Where " &_
									"	( " &_
									"		( " &_
									"				C.ID_Rel_Cadastro = " & RS_Busca("IRC") & " " &_
									"			AND C.ID_Visitante = " & RS_Busca("ID_Visitante") & " " &_
									"		) " &_
									"	OR " &_
									"		(	 " &_
									"				V.CPF = '" & RS_Busca("CPF") & "' " &_
									"			AND	V.Passaporte = '" & RS_Busca("Passaporte") & "' " &_
									"		) " &_
									"	) " &_
									"	AND P.ID_Pedido = " & Pedido & "" &_
									"	AND P.ID_Edicao = " & Session("cliente_edicao") & ""&_
									"	AND C.Cancelado = 0"
			
'RESPONSE.WRITE SQL_Busca_Carrinho
			'RESPONSE.END
			Set RS_Busca_Carrinho = Server.CreateObject("ADODB.Recordset")
			RS_Busca_Carrinho.Open SQL_Busca_Carrinho, Conexao, 3, 3
			
			
			' possui registro
			If Not RS_Busca_Carrinho.Eof Then
				
				Do While Not RS_Busca_Carrinho.EOF 
					' Se tiver o mesmo CPF no Carrinho, abortar
					If Cstr(RS_Busca_Carrinho("CPF")) = Cstr(Form_Busca) OR Cstr(RS_Busca_Carrinho("Passaporte")) = Cstr(Form_Busca) Then
					
						If Cstr(RS_Busca("ID_Visitante")) = Cstr(Session("cliente_visitante")) Then
							Compra = 4
						Else
							Compra = 1
						End If
						
						
						Exit Do
					End If
					RS_Busca_Carrinho.MoveNext
				Loop
				
				If Cint(RS_Busca_Carrinho("ID_Pedido")) = Cint(Pedido) Then
					Compra = 2
				End If
			Else
				Compra = 0			
			End If
			RS_Busca_Carrinho.Close
			
			'Se não existir no meu carrinho, buscar em outros carrinhos		
			If Compra = 0 Then
				' Luiz Ricardo 13/03 - Mudar busca para CPF
				' Busca em todos os carrinhos por pedidos já finalizados ou pendentes de pagamento
				SQL_Busca_Carrinho = 	"Select " &_
										"	C.ID_Carrinho,  " &_
										"	C.ID_Visitante,  " &_
										"	C.ID_Pedido,  " &_
										"	C.ID_Rel_Cadastro, " &_
										"	P.Status_Pedido, " &_
										"	V.Nome_Completo, " &_
										"	V.CPF, " &_
										"	V.Passaporte, " &_
										"	P.ID_Edicao " &_
										"From  Pedidos_Carrinho  As C " &_
										"Inner Join Visitantes As V On V.ID_Visitante = C.ID_Visitante " &_
										"Inner Join Pedidos As P On P.ID_Pedido = C.ID_Pedido " &_
										"Where " &_
										"	( " &_
										"		( " &_
										"				C.ID_Rel_Cadastro = " & RS_Busca("IRC") & " " &_
										"			AND C.ID_Visitante = " & RS_Busca("ID_Visitante") & " " &_
										"		) " &_
										"	OR " &_
										"		(	 " &_
										"				V.CPF = '" & RS_Busca("CPF") & "' " &_
										"			AND	V.Passaporte = '" & RS_Busca("Passaporte") & "' " &_
										"		) " &_
										"	) " &_
										"	AND P.Status_Pedido In (2,3)" &_
										"	AND P.ID_Edicao = " & Session("cliente_edicao")
	
	
	'			SQL_Busca_Carrinho = 	"Select " &_
	'									"	C.ID_Carrinho,  " &_
	'									"	C.ID_Visitante,  " &_
	'									"	C.ID_Pedido,  " &_
	'									"	C.ID_Rel_Cadastro, " &_
	'									"	P.Status_Pedido, " &_
	'									"	V.Nome_Completo " &_
	'									"From  Pedidos_Carrinho  As C " &_
	'									"Inner Join Visitantes As V On V.ID_Visitante = C.ID_Visitante " &_
	'									"Inner Join Pedidos As P On P.ID_Pedido = C.ID_Pedido " &_
	'									"Where " &_
	'									"	(V.CPF = '" & RS_Busca("CPF") & "' " &_
	'									"	Or V.Passaporte = '" & RS_Busca("Passaporte") & "') " &_
	'									"	And P.Status_Pedido In (1,2,3)"
	'
	
				'Response.Write(SQL_Busca_Carrinho & "<br>")
	
				Set RS_Busca_Carrinho = Server.CreateObject("ADODB.Recordset")
				RS_Busca_Carrinho.Open SQL_Busca_Carrinho, Conexao, 3, 3
				
				' possui registro
				If Not RS_Busca_Carrinho.Eof Then
					
					Do While Not RS_Busca_Carrinho.EOF
					
						'Response.Write(Cstr(RS_Busca("ID_Visitante")) = Cstr(Session("cliente_visitante")))
					
						' Se tiver o mesmo CPF no Carrinho, abortar
						If Cstr(RS_Busca_Carrinho("CPF")) = Cstr(Form_Busca) OR Cstr(RS_Busca_Carrinho("Passaporte")) = Cstr(Form_Busca) Then
							If Cstr(RS_Busca("ID_Visitante")) = Cstr(Session("cliente_visitante")) Then
								Compra = 4
							Else
								Compra = 1
							End If
							Exit Do
						End If
						RS_Busca_Carrinho.MoveNext
					Loop
					
					'Response.Write(Compra)
					
					If Cint(RS_Busca_Carrinho("ID_Pedido")) = Cint(Pedido) Then
						Compra = 2
					End If
				Else
					Compra = 0			
				End If
				RS_Busca_Carrinho.Close
			End If
		
			If Compra = 0 Then
				Response.Write(Rs_Busca("ID_Visitante") & ";" & Rs_Busca("IRC") & ";" & Rs_Busca("Nome_Completo") & ";" & Rs_Busca("Data_Nasc"))
			ElseIf Compra = 2 Then
				Response.Write("Erro;2")
			ElseIf Compra = 1 Then
				Response.Write("Erro;1")
			ElseIf Compra = 4 Then
				Response.Write("Erro;4")
			End If
			else
			Response.Write("Erro;5")
			end if
		Else
			Response.Write("Erro;0")
		End If

	Rs_Busca.Close
End If
%>