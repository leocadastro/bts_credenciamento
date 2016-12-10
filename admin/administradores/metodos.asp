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
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================
' METODOS
Select Case Acao
	'==================================================
	Case "add_admin"
		id_perfil = Limpar_Texto(Request("perfil"))
		nome = Limpar_Texto(Request("nome"))
		departamento = Limpar_Texto(Request("departamento"))
		email = Limpar_Texto(Request("email"))
		ativo = Limpar_Texto(Request("ativo"))
		senha = "bts"
		
		SQL_Verificar =	"Select id_admin " &_
						"From Administradores " &_
						"Where nome = '" & nome & "'"
		Set RS_Verificar = Server.CreateObject("ADODB.Recordset")
		RS_Verificar.Open SQL_Verificar, Conexao

		' Se  nao existir insira
		If RS_Verificar.BOF or RS_Verificar.EOF Then
			SQL_Inserir = 	"Insert Into Administradores " &_
							"(id_perfil, nome, departamento, email, senha, ativo) " &_
							"Values " &_
							"('" & id_perfil & "','" & nome & "','" & departamento & "','" & email & "','" & senha & "','" & ativo & "')"
			
			response.write(SQL_Inserir)
			response.write("<br><a href='default.asp'>Voltar</a>")
			
			Set RS_Inserir = Server.CreateObject("ADODB.Recordset")
			RS_Inserir.Open SQL_Inserir, Conexao
			
			response.Redirect("default.asp?msg=add_ok")
		Else
			RS_Verificar.Close
			response.Redirect("default.asp?msg=add_erro_existe")
		End If
	'==================================================
	Case "upd_admin"
		' Campos POST 
		id_perfil = Limpar_Texto(Request("perfil"))
		nome = Limpar_Texto(Request("nome"))
		departamento = Limpar_Texto(Request("departamento"))
		email = Limpar_Texto(Request("email"))
		ativo = Limpar_Texto(Request("ativo"))
		
		nova_senha = Limpar_Texto(Request("nova_senha"))

		If Len(nova_senha) > 0 Then
			SQL_Senha = 	"Update Administradores " &_
							"Set " &_
							"	senha = '" & nova_senha & "' " &_
							"Where id_admin = " & id
						
			Set RS_Senha = Server.CreateObject("ADODB.Recordset")
			RS_Senha.Open SQL_Senha, Conexao
		End If
		
		SQL_Update = 	"Update Administradores " &_
						"Set " &_
						"	id_perfil = '" & id_perfil & "', " &_
						"	nome = '" & nome & "', " &_
						"	departamento = '" & departamento & "', " &_
						"	email = '" & email & "', " &_
						"	ativo = '" & ativo & "' " &_
						"Where id_admin = " & id
		
		response.write("<hr>" & SQL_Update & "<hr>")
		
		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
		
		url = Request("url")
		If Len(url) > 0 Then
			response.Redirect(url & "?msg=upd_ok")
		Else
			response.Redirect("default.asp?msg=upd_ok")
		End If
	'==================================================	
	Case "desativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update Administradores " &_
						"Set " &_
						"	ativo = 0 " &_
						"Where id_admin = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?des_ok")
	'==================================================
	Case "ativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update Administradores " &_
						"Set " &_
						"	ativo = 1 " &_
						"Where id_admin = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?atv_ok")
	'==================================================	
	
End Select

Conexao.Close
%>