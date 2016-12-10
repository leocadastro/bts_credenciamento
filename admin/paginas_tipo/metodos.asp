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
	Case "add_tipo"
		nome	= Limpar_Texto(Request("nome"))
		ativo	= Limpar_Texto(Request("ativo"))
		
		SQL_Verificar =	"Select ID_Tipo " &_
						"From Paginas_Tipo " &_
						"Where Tipo = '" & nome & "'"
		Set RS_Verificar = Server.CreateObject("ADODB.Recordset")
		RS_Verificar.Open SQL_Verificar, Conexao

		' Se  nao existir insira
		If RS_Verificar.BOF or RS_Verificar.EOF Then
			SQL_Inserir = 	"Insert Into Paginas_Tipo " &_
							"(Tipo, Ativo) " &_
							"Values " &_
							"('" & nome & "','" & ativo & "')"
			
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
	Case "upd_tipo"
		' Campos POST 
		nome 	= Limpar_Texto(Request("nome"))
		ativo 	= Limpar_Texto(Request("ativo"))

		SQL_Update = 	"Update Paginas_Tipo " &_
						"Set " &_
						"	Tipo = '" & nome & "', " &_
						"	Ativo = '" & ativo & "' " &_
						"Where ID_Tipo = '" & id & "'"
		
		response.write("<hr>" & SQL_Update & "<hr>")
		
		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
		
		response.Redirect("default.asp?msg=upd_ok")
	'==================================================	
	Case "desativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update Paginas_Tipo " &_
						"Set " &_
						"	Ativo = 0 " &_
						"Where ID_Tipo = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?msg=des_ok")
	'==================================================
	Case "ativar"
		id = Limpar_Texto(Request("id"))
		
		SQL_Update =	"Update Paginas_Tipo " &_
						"Set " &_
						"	Ativo = 1 " &_
						"Where ID_Tipo = " & id

		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
	
		response.Redirect("default.asp?msg=atv_ok")
	'==================================================	
	
End Select

Conexao.Close
%>