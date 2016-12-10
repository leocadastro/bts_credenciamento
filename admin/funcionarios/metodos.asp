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
	Case "add_funcionarios"
		funcionarios_Qtde_PTB	= Limpar_Texto(Request("funcionarios_Qtde_PTB"))
		funcionarios_Qtde_ENG	= Limpar_Texto(Request("funcionarios_Qtde_ENG"))
		funcionarios_Qtde_ESP	= Limpar_Texto(Request("funcionarios_Qtde_ESP"))
		ativo 					= Limpar_Texto(Request("ativo"))

		SQL_Inserir = 	"Insert Into Funcionarios_Qtde " &_
						"(Funcionarios_Qtde_PTB, Funcionarios_Qtde_ENG, Funcionarios_Qtde_ESP, ativo) " &_
						"Values " &_
						"('" & funcionarios_Qtde_PTB & "','" & funcionarios_Qtde_ENG & "','" & funcionarios_Qtde_ESP & "','" & ativo & "')"
		
		Set RS_Inserir = Server.CreateObject("ADODB.Recordset")
		RS_Inserir.Open SQL_Inserir, Conexao
		
		response.Redirect("default.asp?msg=add_ok")
		response.write(SQL_Inserir)
		response.write("<br><a href='default.asp?msg=add_ok'>Voltar</a>")
	'==================================================
	Case "upd_funcionarios"
		' Campos POST 
		funcionarios_Qtde_PTB	= Limpar_Texto(Request("funcionarios_Qtde_PTB"))
		funcionarios_Qtde_ENG	= Limpar_Texto(Request("funcionarios_Qtde_ENG"))
		funcionarios_Qtde_ESP	= Limpar_Texto(Request("funcionarios_Qtde_ESP"))
		ativo 		= Limpar_Texto(Request("ativo"))

		SQL_Update = 	"Update Funcionarios_Qtde " &_
						"Set " &_
						"	Funcionarios_Qtde_PTB = '" & funcionarios_Qtde_PTB & "', " &_
						"	Funcionarios_Qtde_ENG = '" & funcionarios_Qtde_ENG & "', " &_
						"	Funcionarios_Qtde_ESP = '" & funcionarios_Qtde_ESP & "', " &_
						"	ativo = '" & ativo & "' " &_
						"Where ID_Funcionarios_Qtde = " & id
		
		response.write("<hr>" & SQL_Update & "<hr>")
		
		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
		
		response.Redirect("default.asp?msg=upd_ok")
	'==================================================	
	Case "desativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update Funcionarios_Qtde " &_
						"Set " &_
						"	ativo = 0 " &_
						"Where ID_Funcionarios_Qtde = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?msg=des_ok")
	'==================================================
	Case "ativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update Funcionarios_Qtde " &_
						"Set " &_
						"	ativo = 1 " &_
						"Where ID_Funcionarios_Qtde = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?msg=atv_ok")
	'==================================================	
	
End Select

Conexao.Close
%>