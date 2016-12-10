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
	Case "add_Ramo"
		ramo_ptb	= Limpar_Texto(Request("ramo_ptb"))
		ramo_eng	= Limpar_Texto(Request("ramo_eng"))
		ramo_esp	= Limpar_Texto(Request("ramo_esp"))
		Complemento 	= Limpar_Texto(Request("Complemento"))
		ativo 		= Limpar_Texto(Request("ativo"))

		SQL_Inserir = 	"Insert Into RamoeAtividade_V2 " &_
						"(Ramo_Atv_PTB, Ramo_Atv_ENG, Ramo_Atv_ESP, Complemento, ativo, id_admin) " &_
						"Values " &_
						"('" & ramo_ptb & "','" & ramo_eng & "','" & ramo_esp & "','" & atividade & "','" & ativo & "','" & Session("admin_id_usuario") & "')"

		response.write(SQL_Inserir)
		
		Set RS_Inserir = Server.CreateObject("ADODB.Recordset")
		RS_Inserir.Open SQL_Inserir, Conexao
		
		response.Redirect("default.asp?msg=add_ok")
		response.write("<br><a href='default.asp?msg=add_ok'>Voltar</a>")
	'==================================================
	Case "upd_ramo"
		' Campos POST 
		id 			= Limpar_Texto(Request("id"))
		ramo_ptb	= Limpar_Texto(Request("ramo_ptb"))
		ramo_eng	= Limpar_Texto(Request("ramo_eng"))
		ramo_esp	= Limpar_Texto(Request("ramo_esp"))
		complemento = Limpar_Texto(Request("complemento"))
		ativo 		= Limpar_Texto(Request("ativo"))

		SQL_Update = 	"Update RamoeAtividade_V2 " &_
						"Set " &_
						"	Ramo_Atv_PTB = '" & ramo_ptb & "', " &_
						"	Ramo_Atv_ENG = '" & ramo_eng & "', " &_
						"	Ramo_Atv_ESP = '" & ramo_esp & "', " &_
						"	complemento = '" & complemento & "', " &_
						"	ID_Admin = '" & Session("admin_id_usuario") & "', " &_
						"	ativo = '" & ativo & "' " &_
						"Where ID_Ramo_Atividade = " & id
		
		response.write("<hr>" & SQL_Update & "<hr>")
		
		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
		
		response.Redirect("default.asp?msg=upd_ok")
	'==================================================	
	Case "desativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update RamoeAtividade_V2 " &_
						"Set " &_
						"	ativo = 0 " &_
						"Where ID_Ramo_Atividade = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?msg=des_ok")
	'==================================================
	Case "ativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update RamoeAtividade_V2 " &_
						"Set " &_
						"	ativo = 1 " &_
						"Where ID_Ramo_Atividade = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?msg=atv_ok")
	'==================================================	
	
End Select

Conexao.Close
%>