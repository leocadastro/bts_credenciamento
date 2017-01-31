<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include virtual="/admin/inc/limpar_texto.asp"-->

<%
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================

'Valor_Ticket = 65.00
'Valor_Ticket = Application("Valor_Ticket")

SQL_Valor_Ticket = "select * from Edicoes_lote where " &_
											"ID_Edicao = '" & Session("cliente_edicao") & "' " &_
											"and Ativo = 1 and GETDATE() between Data_Inicio and Data_Fim"


Set RS_Consulta_Pedidos = Server.CreateObject("ADODB.Recordset")
RS_Consulta_Pedidos.Open SQL_Valor_Ticket, Conexao, 3, 3

If Not RS_Consulta_Pedidos.Eof Then

	Valor_Ticket = FormatNumber(RS_Consulta_Pedidos("Valor"),2)

Else

	'Valor_Ticket = 70.00
	Response.Write("EDICAO NAO CADASTRADA")
	Response.End

End If

If Limpar_texto(Request("aceito")) = 0 Then

	Response.Redirect("/tickets/status.asp")

Else

	'For Each item In Request.Form
	'	Response.Write "" & item & " - Value: " & Request.Form(item) & "<BR />"
	'Next

	'Response.Write(Session("cliente_edicao") & " <br>" & Session("cliente_idioma") & "<br>" & Session("cliente_logado") & "<br>" & Session("cliente_visitante"))
	'Response.End()


	If Session("cliente_edicao") = "" OR Session("cliente_idioma") = "" OR Session("cliente_logado") = "" or Session("cliente_visitante") = "" Then
		    response.Redirect("http://www.mbxeventos.net/AOLABF2017/")
	End If

	'Response.End()


	ID_Edicao               = Session("cliente_edicao")
	Idioma                  = Session("cliente_idioma")
	ID_TP_Credenciamento    = Session("cliente_tipo")
	TP_Formulario           = Session("cliente_formulario")
	IRC                     = Session("cliente_cadastro")
	ID_Empresa              = Session("cliente_empresa")
	ID_Visitante            = Session("cliente_visitante")
	Nome_Visitante          = Session("cliente_nome")
	CPF_Visitante           = Session("cliente_cpf")

	Acao = Limpar_texto(Request("acao"))

	If Acao = "novo" Then
		NovoPedido()
		Response.Redirect("/tickets/novo_pedido.asp")

	ElseIf Acao = "adicionar" Then
		AdicionarCarrinho Limpar_texto(Request("IdVisitante")),Limpar_texto(Request("IRC")),Limpar_texto(Request("ID_Pedido"))
		Response.Redirect("/tickets/novo_pedido.asp")

	ElseIf Acao = "remover" Then
		RemoverCarrinho Limpar_texto(Request("id")),Limpar_texto(Request("Pedido"))
		Response.Redirect("/tickets/novo_pedido.asp")

	End If

	Function AdicionarCarrinho(Cadastro,Relacionamento,Pedido)
		SQL_Adicionar_Visitante = 	"Insert Into Pedidos_Carrinho " &_
									"	(ID_Visitante,ID_Pedido,ID_Rel_Cadastro) " &_
									"Values " &_
									"	(" & Cadastro & "," & Pedido & "," & Relacionamento & ") "
		'Response.Write(SQL_Adicionar_Visitante)
		Set RS_Adicionar_Visitante = Conexao.Execute(SQL_Adicionar_Visitante)

		AtualizarPedido(Pedido)
	End Function


	Function RemoverCarrinho(IDCarrinho,Pedido)
		SQL_Remover_Visitante = 	"Delete From Pedidos_Carrinho Where ID_Carrinho = " & IDCarrinho
		'Response.Write(SQL_Adicionar_Visitante)
		Set RS_Remover_Visitante = Conexao.Execute(SQL_Remover_Visitante)

		AtualizarPedido(Pedido)
	End Function


	Function AtualizarPedido(Pedido)
		SQL_Quant_Carrinho = 	"Select Count(ID_Carrinho) As Quantidade " &_
								"From Pedidos_Carrinho " &_
								"Where ID_Pedido = " & Pedido & " "&_
								"	And Cancelado = 0"
		'Response.Write(SQL_Quant_Carrinho)
		'Response.End
		Set RS_Quant_Carrinho = Server.CreateObject("ADODB.Recordset")
		RS_Quant_Carrinho.Open SQL_Quant_Carrinho, Conexao, 3, 3

		If Not RS_Quant_Carrinho.Eof Then

		Valor_Atual = Cint(RS_Quant_Carrinho("Quantidade")) * Valor_Ticket
		Valor_Atual = Replace(Valor_Atual, ",", ".")
		Qtde = Cint(RS_Quant_Carrinho("Quantidade"))
		'Response.Write(Valor_Atual)

		'Response.Write(Valor_Atual)
		'Response.End

		End If

		'Response.Write("dsadas: " + ValorAtual)

		SQL_Atualiza_Pedido = 	"Update Pedidos Set " &_
								"	Valor_Pedido = '" & Valor_Atual & "',  " &_
								"	Quantidade = '" & qtde & "'  " &_
								"Where ID_Pedido = " & Pedido

		'Response.Write(SQL_Atualiza_Pedido)
		Set RS_Atualiza_Pedido = Conexao.Execute(SQL_Atualiza_Pedido)

	End Function


	Function NovoPedido()
		SQL_Evento = 	"Select " &_
						"	Ev.ID_Protheus_Evento As Sigla  " &_
						"From Eventos As Ev " &_
						"Inner Join Eventos_Edicoes As Ee " &_
						"	On Ev.ID_Evento = Ee.ID_Evento " &_
						"Where Ee.ID_Edicao = " & ID_Edicao
		'Response.Write(SQL_Evento & "<br>")

		Set RS_Evento = Server.CreateObject("ADODB.Recordset")
		RS_Evento.Open SQL_Evento, Conexao, 3, 3

		If Not RS_Evento.Eof then

			Sigla_Evento = Left(RS_Evento("Sigla"),3)

		Else

			response.Redirect("/?erro=1")

		End If

		SQL_Novo_Pedido = 	"SET NOCOUNT ON;" &_
							" " & vbCrLf & " " &_
							"Insert Into Pedidos " &_
							"	(ID_Edicao,  " &_
							"	ID_Idioma,  " &_
							"	ID_Rel_Cadastro, " &_
							"	ID_Visitante,  " &_
							"	Status_Pedido,  " &_
							"	Valor_Pedido) " &_
							"Values " &_
							"	(" &ID_Edicao  & ",  " &_
							"	" & Idioma & ",  " &_
							"	" & IRC & ", " &_
							"	" & ID_Visitante & ",  " &_
							"	1,  " &_
							"	0) " &_
							" " & vbCrLf & " " &_
							"SELECT @@IDENTITY As NovoID"
		'Response.Write(SQL_Novo_Pedido & "<br>")

		Set RS_Novo_Pedido = Conexao.Execute(SQL_Novo_Pedido)
		NovoID = RS_Novo_Pedido.Fields("NovoID").value

		For I = 1 to (10 - Len(Cstr(NovoID)))
			NumPedido = NumPedido & "0"
		Next

		Numero_Pedido = "E" & Sigla_Evento & Year(Now()) & NumPedido & NovoID
		'Response.Write(Numero_Pedido)

		SQL_Atualiza_Pedido = 	"Update Pedidos Set " &_
								"	Numero_Pedido = '" & Numero_Pedido & "' " &_
								"Where ID_Pedido = " & NovoID
		Set RS_Atualiza_Pedido = Conexao.Execute(SQL_Atualiza_Pedido)



		SQL_Busca_Carrinho = 	"Select " &_
								"	C.ID_Carrinho,  " &_
								"	C.ID_Visitante,  " &_
								"	C.ID_Pedido,  " &_
								"	C.ID_Rel_Cadastro, " &_
								"	P.Status_Pedido, " &_
								"	V.Nome_Completo " &_
								"From  Pedidos_Carrinho  As C " &_
								"Inner Join Visitantes As V On V.ID_Visitante = C.ID_Visitante " &_
								"Inner Join Pedidos As P On P.ID_Pedido = C.ID_Pedido " &_
								"Where " &_
								"	C.ID_Rel_Cadastro = " & IRC & " " &_
								"	And C.ID_Visitante = " & ID_Visitante & " " &_
								"	And P.Status_Pedido In (2,3)" &_
								"	And P.ID_Edicao = " & Session("cliente_edicao")
		Response.Write(SQL_Busca_Carrinho & "<br>")
		Set RS_Busca_Carrinho = Server.CreateObject("ADODB.Recordset")
		RS_Busca_Carrinho.Open SQL_Busca_Carrinho, Conexao, 3, 3

		If RS_Busca_Carrinho.Eof Then
			AdicionarCarrinho ID_Visitante,IRC,NovoID
		End If

		RS_Busca_Carrinho.Close

		Session("Novo_Pedido") = True
		'Response.Write("NOVO PEDIDO GERADO! &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href='/tickets/novo_pedido.asp'>Continuar</a>")
	End Function

End If
%>
