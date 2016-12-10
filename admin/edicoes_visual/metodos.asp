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
	Case "add_configuracao"
		id_edicao 		= Limpar_Texto(Request("id_edicao"))
		cor_fundo 		= Limpar_Texto(Request("cor_fundo"))
		faixa_fundo 	= Limpar_Texto(Request("faixa_fundo"))
		logo_faixa 		= Limpar_Texto(Request("logo_faixa"))
		logo_box 		= Limpar_Texto(Request("logo_box"))
		logo_email		= Limpar_Texto(Request("logo_email"))
		url_template	= Limpar_Texto(Request("url_template"))
		
		SQL_Verificar =	"Select id_edicao " &_
						"From Edicoes_Configuracao " &_
						"Where id_edicao = '" & id_edicao & "'"
		Set RS_Verificar = Server.CreateObject("ADODB.Recordset")
		RS_Verificar.Open SQL_Verificar, Conexao

		' Se  nao existir insira
		If RS_Verificar.BOF or RS_Verificar.EOF Then
			SQL_Inserir = 	"Insert Into Edicoes_Configuracao " &_
							"(ID_Edicao, Cor, Logo_Box, Logo_Negativo, Logo_Email, Faixa_Fundo, Template_Email) " &_
							"Values " &_
							"('" & id_edicao & "','" & cor_fundo & "','" & logo_box & "','" & logo_faixa & "','" & logo_email & "','" & faixa_fundo & "','" & url_template & "')"
			
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
	Case "upd_visual"
		' Campos POST 
		id_edicao = Limpar_Texto(Request("id_edicao"))
		cor_fundo = Limpar_Texto(Request("cor_fundo"))
		faixa_fundo = Limpar_Texto(Request("faixa_fundo"))
		logo_faixa = Limpar_Texto(Request("logo_faixa"))
		logo_box = Limpar_Texto(Request("logo_box"))
		logo_email		= Limpar_Texto(Request("logo_email"))

		SQL_Update = 	"Update Edicoes_Configuracao " &_
						"Set " &_
						"	ID_Edicao = '" & id_edicao & "', " &_
						"	Cor = '" & cor_fundo & "', " &_
						"	Logo_Box = '" & logo_box & "', " &_
						"	Logo_Negativo = '" & logo_faixa & "', " &_
						"	Logo_Email = '" & logo_email & "', " &_
						"	Faixa_Fundo = '" & faixa_fundo & "', " &_
						"	Template_Email = '" & url_template & "' " &_
						"Where ID_Configuracao = " & id
		
		response.write("<hr>" & SQL_Update & "<hr>")
		
		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
		
		response.Redirect("default.asp?msg=upd_ok")
	'==================================================	
	Case "desativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update Edicoes_Configuracao " &_
						"Set " &_
						"	ativo = 0 " &_
						"Where ID_Configuracao = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?des_ok")
	'==================================================
	Case "ativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update Edicoes_Configuracao " &_
						"Set " &_
						"	ativo = 1 " &_
						"Where ID_Configuracao = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?atv_ok")
	'==================================================	
	
End Select

Conexao.Close
%>