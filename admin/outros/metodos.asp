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
	Case "atualizar"
		' Campos POST 
		id 			= Limpar_Texto(Request("id"))
		registro 	= Limpar_Texto(Request("registro"))
		tabela	= Limpar_Texto(Request("tab"))
		
		If tabela = "ramo" then
			campo_tabela = "Ramo"
		ElseIf tabela = "atividade" then
			campo_tabela = "Atividade"
		End If

		SQL_Update = 	"Update Relacionamento_" & campo_tabela & " " &_
						"Set " &_
						"	" & campo_tabela & "_Outros = '" & registro & "' " &_
						"Where ID_Relacionamento_" & campo_tabela & " = " & id
		
		response.write("<hr>" & SQL_Update & "<hr>")
		
		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
		
		response.Redirect("listar.asp?tab=" &tabela& "&msg=upd_ok")
		
	'==================================================
	Case "atualizar_subcargo"
		' Campos POST 
		id 			= Limpar_Texto(Request("id"))
		registro 	= Limpar_Texto(Request("registro"))


		SQL_Update = 	"Update Visitantes " &_
						"Set " &_
						"	SubCargo_Outros = '" & registro & "' " &_
						"Where ID_Visitante = " & id
		
		response.write("<hr>" & SQL_Update & "<hr>")
		
		Set RS_Update = Server.CreateObject("ADODB.Recordset")
		RS_Update.Open SQL_Update, Conexao
		
		response.Redirect("subcargos.asp?msg=upd_ok")
	
End Select

Conexao.Close
%>