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
	Case "add_interesse"
		InteresseFeira_ptb	= Limpar_Texto(Request("InteresseFeira_ptb"))
		InteresseFeira_eng	= Limpar_Texto(Request("InteresseFeira_eng"))
		InteresseFeira_esp	= Limpar_Texto(Request("InteresseFeira_esp"))
		ativo 		= Limpar_Texto(Request("ativo"))

		SQL_Inserir = 	"Insert Into ProdutoFeira " &_
						"(Feira_ptb, Feira_eng, Feira_esp, ativo) " &_
						"Values " &_
						"('" & InteresseFeira_ptb & "','" & InteresseFeira_eng & "','" & InteresseFeira_esp & "','" & ativo & "')"
		
		Set RS_Inserir = Server.CreateObject("ADODB.Recordset")
		RS_Inserir.Open SQL_Inserir, Conexao
		
		response.Redirect("default.asp?msg=add_ok")
		response.write(SQL_Inserir)
		response.write("<br><a href='default.asp?msg=add_ok'>Voltar</a>")
	'==================================================
	Case "upd_interesse"
		' Campos POST 
		id 			= Limpar_Texto(Request("id"))
		InteresseFeira_ptb	= Limpar_Texto(Request("InteresseFeira_ptb"))
		InteresseFeira_eng	= Limpar_Texto(Request("InteresseFeira_eng"))
		InteresseFeira_esp	= Limpar_Texto(Request("InteresseFeira_esp"))
		ativo 		= Limpar_Texto(Request("ativo"))

		SQL_Update = 	"Update ProdutoFeira " &_
						"Set " &_
						"	Feira_ptb = '" & InteresseFeira_ptb & "', " &_
						"	Feira_eng = '" & InteresseFeira_eng & "', " &_
						"	Feira_esp = '" & InteresseFeira_esp & "', " &_
						"	ativo = '" & ativo & "' " &_
						"Where id_Feira = " & id
		
		response.write("<hr>" & SQL_Update & "<hr>")
		
		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
		
		response.Redirect("default.asp?msg=upd_ok")
	'==================================================	
	Case "desativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update ProdutoFeira " &_
						"Set " &_
						"	ativo = 0 " &_
						"Where id_Feira = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?msg=des_ok")
	'==================================================
	Case "ativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update ProdutoFeira " &_
						"Set " &_
						"	ativo = 1 " &_
						"Where id_Feira = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?msg=atv_ok")
	'==================================================	
	
End Select

Conexao.Close
%>