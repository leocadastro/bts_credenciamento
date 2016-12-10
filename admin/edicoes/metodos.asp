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
	Case "add_edicao"
		id_evento 	= Limpar_Texto(Request("evento"))
		ano 		= Limpar_Texto(Request("ano"))
		hora_ini	= Limpar_Texto(Request("hora_ini"))
		data_ini	= Limpar_Texto(Request("data_ini"))
		hora_fim	= Limpar_Texto(Request("hora_fim"))
		data_fim	= Limpar_Texto(Request("data_fim"))
		ativo 		= Limpar_Texto(Request("ativo"))

		dia = Left(data_ini, 2)
		mes = Mid(data_ini, 4, 2)
		ano = Right(data_ini, 4)
		inicio 	= "'" & ano & "-" & mes & "-" & dia & " " & hora_ini & ":01.000'"

		diaf = Left(data_fim, 2)
		mesf = Mid(data_fim, 4, 2)
		anof = Right(data_fim, 4)
		fim 	= "'" & anof & "-" & mesf & "-" & diaf & " " & hora_fim & ":01.000'"
		
		SQL_Verificar =	"Select id_evento " &_
						"From Eventos_Edicoes " &_
						"Where id_evento = '" & id_evento & "' and ano = " & ano
		Set RS_Verificar = Server.CreateObject("ADODB.Recordset")
		RS_Verificar.Open SQL_Verificar, Conexao

		' Se  nao existir insira
		If RS_Verificar.BOF or RS_Verificar.EOF Then
			SQL_Inserir = 	"Insert Into Eventos_Edicoes " &_
							"(id_evento, ano, data_inicio_feira, data_fim_feira, ativo) " &_
							"Values " &_
							"(" & id_evento & "," & ano & "," & inicio & "," & fim & "," & ativo & ")"

							response.write(SQL_Inserir)
			
			Set RS_Inserir = Server.CreateObject("ADODB.Recordset")
			RS_Inserir.Open SQL_Inserir, Conexao
			
			response.Redirect("default.asp?msg=add_ok")
			response.write(SQL_Inserir)
			response.write("<br><a href='default.asp?msg=add_ok'>Voltar</a>")
		Else
			RS_Verificar.Close
			response.Redirect("default.asp?msg=add_erro_existe")
		End If
	'==================================================
	Case "upd_edicao"
		' Campos POST 
		id_evento 	= Limpar_Texto(Request("evento"))
		ano 		= Limpar_Texto(Request("ano"))
		data_ini	= Limpar_Texto(Request("data_ini"))
		hora_ini	= Limpar_Texto(Request("hora_ini"))
		hora_fim	= Limpar_Texto(Request("hora_fim"))
		data_fim	= Limpar_Texto(Request("data_fim"))
		ativo 		= Limpar_Texto(Request("ativo"))

		dia = Left(data_ini, 2)
		mes = Mid(data_ini, 4, 2)
		ano = Right(data_ini, 4)
		inicio 	= "'" & ano & "-" & mes & "-" & dia & " " & hora_ini & ":01.000'"

		diaf = Left(data_fim, 2)
		mesf = Mid(data_fim, 4, 2)
		anof = Right(data_fim, 4)
		fim 	= "'" & anof & "-" & mesf & "-" & diaf & " " & hora_fim & ":01.000'"

		SQL_Update = 	"Update Eventos_Edicoes " &_
						"Set " &_
						"	id_evento = '" & id_evento & "', " &_
						"	ano = '" & ano & "', " &_
						"	data_inicio_feira = " & inicio & ", " &_
						"	data_fim_feira = " & fim & ", " &_
						"	ativo = '" & ativo & "' " &_
						"Where id_edicao = " & id
		
		response.write("<hr>" & SQL_Update & "<hr>")
		
		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
		
		response.Redirect("default.asp?msg=upd_ok")
	'==================================================	
	Case "desativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update Eventos_Edicoes " &_
						"Set " &_
						"	ativo = 0 " &_
						"Where id_edicao = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?msg=des_ok")
	'==================================================
	Case "ativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update Eventos_Edicoes " &_
						"Set " &_
						"	ativo = 1 " &_
						"Where id_edicao = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?msg=atv_ok")
	'==================================================	
	
End Select

Conexao.Close
%>