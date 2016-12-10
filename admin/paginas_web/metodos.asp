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
	Case "add_pagina"
		pagina = Limpar_Texto(Request("pagina"))
			
		SQL_Verificar =	"Select id_pagina " &_
						"From Paginas_Web " &_
						"Where " &_
						"	pagina = '" & pagina & "' "
						
		Set RS_Verificar = Server.CreateObject("ADODB.Recordset")
		RS_Verificar.Open SQL_Verificar, Conexao

		' Se  nao existir insira
		If RS_Verificar.BOF or RS_Verificar.EOF Then
			SQL_Inserir = 	"Insert Into Paginas_Web " &_
							"(pagina, ID_Admin) " &_
							"Values " &_
							"('" & pagina & "'," & Session("admin_id_usuario") & ")"
			
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
	Case "upd_pagina"
		' Campos POST 
		pagina = Limpar_Texto(Request("pagina"))

		SQL_Update = 	"Update Paginas_Web " &_
						"Set " &_
						"	pagina = '" & pagina & "' " &_
						"Where id_pagina = " & id
		
		response.write("<hr>" & SQL_Update & "<hr>")
		
		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
		
		response.Redirect("default.asp?msg=upd_ok")
	'==================================================	
	
End Select

Conexao.Close
%>