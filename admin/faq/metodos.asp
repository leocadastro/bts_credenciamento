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
	Case "atualizar"
	
		descricao 	= Limpar_Texto(Request("descricao"))
		
		SQL_Verificar =	"Update FAQ " &_
						"Set Descricao = '" & descricao & "' " &_
						"Where  " &_
						"	ID_FAQ = 1" &_
						
		response.write(SQL_Verificar)
		
		Set RS_Verificar = Server.CreateObject("ADODB.Recordset")
		RS_Verificar.Open SQL_Verificar, Conexao

			
			response.write("<br><a href='default.asp?msg=add_ok'>Voltar</a>")
			response.Redirect("default.asp?msg=add_ok")
	'==================================================

	Case "novo"
	
		SQL_Verificar =	"Insert Into FAQ " &_
						"	(Descricao)	Values ('Primeiro Teste')" &_
						
		response.write(SQL_Verificar)
		
		Set RS_Verificar = Server.CreateObject("ADODB.Recordset")
		RS_Verificar.Open SQL_Verificar, Conexao
		
End Select

Conexao.Close
%>