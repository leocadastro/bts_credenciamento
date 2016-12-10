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
	Case "add_cargo"
		cargo_ptb	= Limpar_Texto(Request("cargo_ptb"))
		cargo_eng	= Limpar_Texto(Request("cargo_eng"))
		cargo_esp	= Limpar_Texto(Request("cargo_esp"))
		subcargo	= Limpar_Texto(Request("subcargo"))
		ativo 		= Limpar_Texto(Request("ativo"))

		SQL_Inserir = 	"Insert Into Cargo " &_
						"(cargo_ptb, cargo_eng, cargo_esp, subcargo, ativo) " &_
						"Values " &_
						"('" & cargo_ptb & "','" & cargo_eng & "','" & cargo_esp & "','" & subcargo & "','" & ativo & "')"
		
		Set RS_Inserir = Server.CreateObject("ADODB.Recordset")
		RS_Inserir.Open SQL_Inserir, Conexao
		
		response.Redirect("default.asp?msg=add_ok")
		response.write(SQL_Inserir)
		response.write("<br><a href='default.asp?msg=add_ok'>Voltar</a>")
	'==================================================
	Case "upd_cargo"
		' Campos POST 
		id_cargo 	= Limpar_Texto(Request("id_cargo"))
		cargo_ptb	= Limpar_Texto(Request("cargo_ptb"))
		cargo_eng	= Limpar_Texto(Request("cargo_eng"))
		cargo_esp	= Limpar_Texto(Request("cargo_esp"))
		subcargo	= Limpar_Texto(Request("subcargo"))
		ativo 		= Limpar_Texto(Request("ativo"))

		SQL_Update = 	"Update Cargo " &_
						"Set " &_
						"	cargo_ptb = '" & cargo_ptb & "', " &_
						"	cargo_eng = '" & cargo_eng & "', " &_
						"	cargo_esp = '" & cargo_esp & "', " &_
						"	subcargo = '" & subcargo & "', " &_
						"	ativo = '" & ativo & "' " &_
						"Where id_cargo = " & id
		
		response.write("<hr>" & SQL_Update & "<hr>")
		
		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
		
		response.Redirect("default.asp?msg=upd_ok")
	'==================================================	
	Case "desativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update Cargo " &_
						"Set " &_
						"	ativo = 0 " &_
						"Where id_cargo = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?msg=des_ok")
	'==================================================
	Case "ativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update Cargo " &_
						"Set " &_
						"	ativo = 1 " &_
						"Where id_cargo = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?msg=atv_ok")
	'==================================================	
	
End Select

Conexao.Close
%>