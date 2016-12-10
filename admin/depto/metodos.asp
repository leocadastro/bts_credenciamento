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
	Case "add_Depto"
		depto_ptb	= Limpar_Texto(Request("depto_ptb"))
		depto_eng	= Limpar_Texto(Request("depto_eng"))
		depto_esp	= Limpar_Texto(Request("depto_esp"))
		subdepto	= Limpar_Texto(Request("subdepto"))
		ativo 		= Limpar_Texto(Request("ativo"))

		SQL_Inserir = 	"Insert Into depto " &_
						"(depto_ptb, depto_eng, depto_esp, ativo) " &_
						"Values " &_
						"('" & depto_ptb & "','" & depto_eng & "','" & depto_esp & "','" & ativo & "')"
		
		Set RS_Inserir = Server.CreateObject("ADODB.Recordset")
		RS_Inserir.Open SQL_Inserir, Conexao
		
		response.Redirect("default.asp?msg=add_ok")
		response.write(SQL_Inserir)
		response.write("<br><a href='default.asp?msg=add_ok'>Voltar</a>")
	'==================================================
	Case "upd_Depto"
		' Campos POST 
		id_depto 	= Limpar_Texto(Request("id_depto"))
		depto_ptb	= Limpar_Texto(Request("depto_ptb"))
		depto_eng	= Limpar_Texto(Request("depto_eng"))
		depto_esp	= Limpar_Texto(Request("depto_esp"))
		ativo 		= Limpar_Texto(Request("ativo"))

		SQL_Update = 	"Update depto " &_
						"Set " &_
						"	depto_ptb = '" & depto_ptb & "', " &_
						"	depto_eng = '" & depto_eng & "', " &_
						"	depto_esp = '" & depto_esp & "', " &_
						"	ativo = '" & ativo & "' " &_
						"Where id_depto = " & id
		
		response.write("<hr>" & SQL_Update & "<hr>")
		
		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
		
		response.Redirect("default.asp?msg=upd_ok")
	'==================================================	
	Case "desativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update depto " &_
						"Set " &_
						"	ativo = 0 " &_
						"Where id_depto = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?msg=des_ok")
	'==================================================
	Case "ativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update depto " &_
						"Set " &_
						"	ativo = 1 " &_
						"Where id_depto = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?msg=atv_ok")
	'==================================================	
	
End Select

Conexao.Close
%>