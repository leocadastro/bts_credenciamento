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
If Session("admin_id_usuario") = "" Then response.Redirect("default.asp?msg=ERRO")
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================
' METODOS
Select Case Acao
	'==================================================
	Case "add_rel_evt_tipo"
		id_edicao 	= Limpar_Texto(Request("id_edicao"))
		id_tipo 	= Limpar_Texto(Request("id_tipo"))
		data_ini	= Limpar_Texto(Request("data_ini"))
		data_fim	= Limpar_Texto(Request("data_fim"))
		hora_ini	= Limpar_Texto(Request("hora_ini"))
		hora_fim	= Limpar_Texto(Request("hora_fim"))
		url			= Limpar_Texto(Request("url"))
		If Len(Trim(url)) = 0 Then url = "null" else url = "'" & url & "'"
		ativo 		= Limpar_Texto(Request("ativo"))
		
		dia = Left(data_ini, 2)
		mes = Mid(data_ini, 4, 2)
		ano = Right(data_ini, 4)
		inicio = "'" & ano & "-" & mes & "-" & dia & " " & hora_ini & ":01.000'"
		
		dia = Left(data_fim, 2)
		mes = Mid(data_fim, 4, 2)
		ano = Right(data_fim, 4)
		fim  = "'" & ano & "-" & mes & "-" & dia & " " & hora_fim & ":01.000'"
		
		SQL_Verificar =	"Select id_edicao_tipo " &_
						"From Edicoes_Tipo " &_
						"Where  " &_
						"	id_edicao = '" & id_edicao & "'  " &_
						"	AND id_tipo_credenciamento = '" & id_tipo & "'"
						
		response.write(SQL_Verificar)
		Set RS_Verificar = Server.CreateObject("ADODB.Recordset")
		RS_Verificar.Open SQL_Verificar, Conexao

		' Se  nao existir insira
		If RS_Verificar.BOF or RS_Verificar.EOF Then
			SQL_Inserir = 	"Insert Into Edicoes_Tipo " &_
							"(ID_Edicao, ID_Tipo_Credenciamento, ID_Admin, Inicio, Fim, URL_Especial, Ativo) " &_
							"Values " &_
							"('" & id_edicao & "','" & id_tipo & "','" & Session("admin_id_usuario") & "'," & inicio & "," & fim & "," & url & ",'" & ativo & "')"
			
			response.write(SQL_Inserir)
			
			Set RS_Inserir = Server.CreateObject("ADODB.Recordset")
			RS_Inserir.Open SQL_Inserir, Conexao
			
			response.write("<br><a href='default.asp?msg=add_ok'>Voltar</a>")
			'response.Redirect("default.asp?msg=add_ok")
		Else
			RS_Verificar.Close
			'response.Redirect("default.asp?msg=add_erro_existe")
		End If
	'==================================================
	Case "upd_rel_evt_tipo"
		' Campos POST 

		id_edicao 	= Limpar_Texto(Request("id_edicao"))
		id_tipo 	= Limpar_Texto(Request("id_tipo"))
		data_ini	= Limpar_Texto(Request("data_ini"))
		data_fim	= Limpar_Texto(Request("data_fim"))
		hora_ini	= Limpar_Texto(Request("hora_ini"))
		hora_fim	= Limpar_Texto(Request("hora_fim"))
		url			= Limpar_Texto(Request("URL_Especial"))
		
		If Len(Trim(url)) = 0 Then 
			url = "null" 
		Else 
			url = "'" & url & "'"
		End If
		ativo 		= Limpar_Texto(Request("ativo"))
		
		dia = Left(data_ini, 2)
		mes = Mid(data_ini, 4, 2)
		ano = Right(data_ini, 4)
		inicio = "'" & ano & "-" & mes & "-" & dia & " " & hora_ini & ":01.000'"
		
		dia = Left(data_fim, 2)
		mes = Mid(data_fim, 4, 2)
		ano = Right(data_fim, 4)
		fim  = "'" & ano & "-" & mes & "-" & dia & " " & hora_fim & ":01.000'"

		SQL_Update = 	"Update Edicoes_Tipo " &_
						"Set " &_
						"	ID_Edicao = '" & ID_Edicao & "', " &_
						"	ID_Tipo_Credenciamento = '" & ID_Tipo & "', " &_
						"	ID_Admin = '" & Session("admin_id_usuario") & "', " &_
						"	Inicio = " & inicio & ", " &_
						"	Fim = " & fim & ", " &_
						"	URL_Especial = " & url & ", " &_
						"	Ativo = '" & ativo & "' " &_
						"Where ID_Edicao_Tipo = " & id
		
		response.write("<hr>" & SQL_Update & "<hr>")
		
		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
		
		response.Redirect("default.asp?msg=upd_ok")
	'==================================================	
	Case "desativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update Edicoes_Tipo " &_
						"Set " &_
						"	ativo = 0 " &_
						"Where ID_Edicao_Tipo = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?des_ok")
	'==================================================
	Case "ativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update Edicoes_Tipo " &_
						"Set " &_
						"	ativo = 1 " &_
						"Where ID_Edicao_Tipo = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?atv_ok")
	'==================================================	
	
End Select

Conexao.Close
%>