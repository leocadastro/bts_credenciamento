<% Response.Expires = -1
Response.AddHeader "Cache-Control", "no-cache"
Response.AddHeader "Pragma", "no-cache" 
%>
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/acentos_htm.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
Id = Limpar_Texto(Request("id"))
Acao = Limpar_Texto(Request("acao"))

'response.write("ID: " & ID)
'response.write("acao: " & Acao)
'response.end()

'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================
' METODOS
Select Case Acao
	'==================================================
	Case "add_pergunta"
		id_formulario	= Limpar_Texto(Request("id_formulario"))
		pergunta_ptb	= Limpar_Texto(Request("pergunta_ptb"))
		pergunta_eng	= Limpar_Texto(Request("pergunta_eng"))
		pergunta_esp	= Limpar_Texto(Request("pergunta_esp"))
		'nome			= Limpar_Texto(Request("nome"))
		tipo			= Limpar_Texto(Request("tipo"))
		multiplo		= Limpar_Texto(Request("multiplo"))
		ativo 			= Limpar_Texto(Request("ativo"))

		SQL_Inserir = 	"Insert Into Perguntas " &_
						"(id_formulario, pergunta_ptb, pergunta_eng, pergunta_esp, tipo, multiplo, ativo) " &_
						"Values " &_
						"('" & id_formulario & "','" & pergunta_ptb & "','" & pergunta_eng & "','" & pergunta_esp & "','" & tipo & "','" & multiplo & "','" & ativo & "')"
		
		response.write(SQL_Inserir)
		
		Set RS_Inserir = Server.CreateObject("ADODB.Recordset")
		RS_Inserir.Open SQL_Inserir, Conexao
		
		response.Redirect("default.asp?msg=add_ok")
		response.write("<br><a href='default.asp?msg=add_ok'>Voltar</a>")
	'==================================================
	Case "upd_pergunta"
		' Campos POST 
		id_pergunta		= Limpar_Texto(Request("id_pergunta"))
		pergunta_ptb	= Limpar_Texto(Request("pergunta_ptb"))
		pergunta_eng	= Limpar_Texto(Request("pergunta_eng"))
		pergunta_esp	= Limpar_Texto(Request("pergunta_esp"))
		tipo			= Limpar_Texto(Request("tipo"))
		multiplo		= Limpar_Texto(Request("multiplo"))

		SQL_Update = 	"Update Perguntas " &_
						"Set " &_
						"	pergunta_ptb = '" & pergunta_ptb & "', " &_
						"	pergunta_eng = '" & pergunta_eng & "', " &_
						"	pergunta_esp = '" & pergunta_esp & "', " &_
						"	tipo = '" & tipo & "', " &_
						"	multiplo = '" & multiplo & "' " &_
						"Where id_perguntas = " & id
		
		response.write("<hr>" & SQL_Update & "<hr>")
		
		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
		
		response.Redirect("default.asp?msg=upd_ok")
	'==================================================	
	Case "desativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update Perguntas " &_
						"Set " &_
						"	ativo = 0 " &_
						"Where id_perguntas = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?msg=des_ok")
	'==================================================
	Case "ativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update Perguntas " &_
						"Set " &_
						"	ativo = 1 " &_
						"Where id_perguntas = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?msg=atv_ok")

	'==================================================	
	Case "add_opcoes"
		id_pergunta	= Limpar_Texto(Request("id_pergunta"))
		opcoes_ptb	= Limpar_Texto(Request("opcoes_ptb"))
		opcoes_eng	= Limpar_Texto(Request("opcoes_eng"))
		opcoes_esp	= Limpar_Texto(Request("opcoes_esp"))
		ordem		= Limpar_Texto(Request("ordem"))
		ativo		= Limpar_Texto(Request("ativo"))

		SQL_Inserir = 	"Insert Into Perguntas_Opcoes " &_
						"(id_perguntas, Opcao_ptb, Opcao_eng, Opcao_esp, Ordem, Ativo) " &_
						"Values " &_
						"('" & id_pergunta & "','" & opcoes_ptb & "','" & opcoes_eng & "','" & opcoes_esp & "','" & ordem & "','" & ativo & "')"
		
		response.write(SQL_Inserir)
		
		Set RS_Inserir = Server.CreateObject("ADODB.Recordset")
		RS_Inserir.Open SQL_Inserir, Conexao
		
		response.Redirect("editar.asp?id=" & id_pergunta & "&msg=add_ok")
	'==================================================
	Case "upd_opcoes"
		' Campos POST 
		id_op 		= Limpar_Texto(Request("id_op"))
		id_pergunta	= Limpar_Texto(Request("id_pergunta"))
		opcoes_ptb	= Limpar_Texto(Request("opcoes_ptb"))
		opcoes_eng	= Limpar_Texto(Request("opcoes_eng"))
		opcoes_esp	= Limpar_Texto(Request("opcoes_esp"))
		ordem		= Limpar_Texto(Request("ordem"))

		SQL_Update = 	"Update Perguntas_Opcoes " &_
						"Set " &_
						"	opcao_ptb = '" & opcoes_ptb & "', " &_
						"	opcao_eng = '" & opcoes_eng & "', " &_
						"	opcao_esp = '" & opcoes_esp & "', " &_
						"	ordem = '" & ordem & "' " &_
						"Where id_opcoes = " & id_op
		
		response.write("<hr>" & SQL_Update & "<hr>")
		
		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
		
		response.Redirect("editar.asp?id=" & id_pergunta & "&msg=upd_ok")
	'==================================================	
	Case "desativar_opcoes"
		id_op = Limpar_Texto(Request("id_op"))
		
		SQL_Update =	"Update Perguntas_Opcoes " &_
						"Set " &_
						"	ativo = 0 " &_
						"Where id_opcoes = " & id_op

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("editar.asp?id=" & id & "&msg=des_ok")
	'==================================================
	Case "ativar_opcoes"
		id_op = Limpar_Texto(Request("id_op"))
		
		SQL_Update =	"Update Perguntas_Opcoes " &_
						"Set " &_
						"	ativo = 1 " &_
						"Where id_opcoes = " & id_op

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("editar.asp?id=" & id & "&msg=atv_ok")
	'==================================================	

	
End Select

Conexao.Close
%>