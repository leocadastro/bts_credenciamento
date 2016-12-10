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
	Case "add_atividade"
		atividade_ptb	= Limpar_Texto(Request("atividade_ptb"))
		atividade_eng	= Limpar_Texto(Request("atividade_eng"))
		atividade_esp	= Limpar_Texto(Request("atividade_esp"))
		ativo 		= Limpar_Texto(Request("ativo"))

		SQL_Inserir = 	"Insert Into AtividadeEconomica " &_
						"(atividade_ptb, atividade_eng, atividade_esp, ativo) " &_
						"Values " &_
						"('" & atividade_ptb & "','" & atividade_eng & "','" & atividade_esp & "','" & ativo & "')"
		
		Set RS_Inserir = Server.CreateObject("ADODB.Recordset")
		RS_Inserir.Open SQL_Inserir, Conexao
		
		response.Redirect("default.asp?msg=add_ok")
		response.write(SQL_Inserir)
		response.write("<br><a href='default.asp?msg=add_ok'>Voltar</a>")
	'==================================================
	Case "upd_atividade"
		' Campos POST 
		id 			= Limpar_Texto(Request("id"))
		atividade_ptb	= Limpar_Texto(Request("atividade_ptb"))
		atividade_eng	= Limpar_Texto(Request("atividade_eng"))
		atividade_esp	= Limpar_Texto(Request("atividade_esp"))
		ativo 		= Limpar_Texto(Request("ativo"))

		SQL_Update = 	"Update AtividadeEconomica " &_
						"Set " &_
						"	atividade_ptb = '" & atividade_ptb & "', " &_
						"	atividade_eng = '" & atividade_eng & "', " &_
						"	atividade_esp = '" & atividade_esp & "', " &_
						"	ativo = '" & ativo & "' " &_
						"Where id_atividade = " & id
		
		response.write("<hr>" & SQL_Update & "<hr>")
		
		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
		
		response.Redirect("default.asp?msg=upd_ok")
	'==================================================	
	Case "desativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update AtividadeEconomica " &_
						"Set " &_
						"	ativo = 0 " &_
						"Where id_atividade = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?msg=des_ok")
	'==================================================
	Case "ativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update AtividadeEconomica " &_
						"Set " &_
						"	ativo = 1 " &_
						"Where id_atividade = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?msg=atv_ok")
	'==================================================	
	
End Select

Conexao.Close
%>