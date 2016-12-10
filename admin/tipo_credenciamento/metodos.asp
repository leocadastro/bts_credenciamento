<% Response.Expires = -1
Response.AddHeader "Cache-Control", "no-cache"
Response.AddHeader "Pragma", "no-cache" 
%>
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/acentos_htm.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%

	For Each item In Request.Form
		Response.Write "Key: " & item & " - Value: " & Request.Form(item) & "<BR />"
	Next

Id = Limpar_Texto(Request("id"))
Acao = Limpar_Texto(Request("acao"))
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================
' METODOS
Select Case Acao
	'==================================================
	Case "add_tipo_cred"
		id_idioma 	= Limpar_Texto(Request("id_idioma"))
		nome 		= Limpar_Texto(Request("nome"))
		img_faixa	= Limpar_Texto(Request("img_faixa"))
		img_box 	= Limpar_Texto(Request("img_box"))
		url			= Limpar_Texto(Request("url"))
		ativo 		= Limpar_Texto(Request("ativo"))
		
		SQL_Verificar =	"Select id_idioma " &_
						"From Tipo_Credenciamento " &_
						"Where id_idioma = '" & id_idioma & "' AND nome = '" & nome & "'"
		Set RS_Verificar = Server.CreateObject("ADODB.Recordset")
		RS_Verificar.Open SQL_Verificar, Conexao

		' Se  nao existir insira
		If RS_Verificar.BOF or RS_Verificar.EOF Then
			SQL_Inserir = 	"Insert Into Tipo_Credenciamento " &_
							"(ID_Idioma, Nome, IMG_Faixa, IMG_Box, URL, Ativo) " &_
							"Values " &_
							"('" & id_idioma & "','" & nome & "','" & img_faixa & "','" & img_box & "','" & url & "','" & ativo & "')"
			
			response.write(SQL_Inserir)
			
			Set RS_Inserir = Server.CreateObject("ADODB.Recordset")
			RS_Inserir.Open SQL_Inserir, Conexao
			
			response.write("<br><a href='default.asp?msg=add_ok'>Voltar</a>")
			response.Redirect("default.asp?msg=add_ok")
		Else
			RS_Verificar.Close
			response.Redirect("default.asp?msg=add_erro_existe")
		End If
	'==================================================
	Case "upd_tipo_cred"
		' Campos POST 
		id_idioma 	= Limpar_Texto(Request("id_idioma"))
		nome 		= Limpar_Texto(Request("nome"))
		img_faixa	= Limpar_Texto(Request("img_faixa"))
		img_box 	= Limpar_Texto(Request("img_box"))
		url			= Limpar_Texto(Request("url"))
		ativo 		= Limpar_Texto(Request("ativo"))

		SQL_Update = 	"Update Tipo_Credenciamento " &_
						"Set " &_
						"	ID_Idioma = '" & ID_Idioma & "', " &_
						"	Nome = '" & nome & "', " &_
						"	IMG_Faixa = '" & img_faixa & "', " &_
						"	IMG_Box = '" & img_box & "', " &_
						"	URL = '" & URL & "', " &_
						"	Ativo = '" & ativo & "' " &_
						"Where ID_Tipo_Credenciamento = " & id
		
		response.write("<hr>" & SQL_Update & "<hr>")
		
		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
		
		response.Redirect("default.asp?msg=upd_ok")
	'==================================================	
	Case "desativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update Tipo_Credenciamento " &_
						"Set " &_
						"	ativo = 0 " &_
						"Where ID_Tipo_Credenciamento = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?des_ok")
	'==================================================
	Case "ativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update Tipo_Credenciamento " &_
						"Set " &_
						"	ativo = 1 " &_
						"Where ID_Tipo_Credenciamento = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?atv_ok")
	'==================================================	
	
End Select

Conexao.Close
%>