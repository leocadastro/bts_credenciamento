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
If Session("admin_id_usuario") = "" Then response.Redirect("default.asp?msg=ERRO")
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================
' METODOS
Select Case Acao
	'==================================================
	Case "menus_relacionar"
		area		= Split( Trim(Request("ID_InteresseFeira")) )
		id_edicao	= Request("id_edicao")
		
		'1 - Remover todos Atualmente Relacionados
		SQL_Del_1 =	"Delete From Relacionamento_Edicoes_InteresseFeira " &_
					"Where " &_
					"	id_edicao = " & id_edicao
		
		Set RS_Del_1 = Server.CreateObject("ADODB.Recordset")
		RS_Del_1.Open SQL_Del_1, Conexao
	
		'3 - Loop para Relacionar cada Ramo
		For i = 0 to Ubound(area)
			SQL_Inserir =	"Insert Into Relacionamento_Edicoes_InteresseFeira " &_
							"(id_edicao, ID_InteresseFeira, id_admin) " &_
							"Values " &_
							"(" & id_edicao & "," & Replace(area(i),",","") & "," & Session("admin_id_usuario") & ")"

			Set RS_Inserir = Server.CreateObject("ADODB.Recordset")
			
'			response.write(SQL_Inserir & "<hr>")
			
			RS_Inserir.Open SQL_Inserir, Conexao
		Next
		response.Redirect("default.asp?msg=add_ok")
	'==================================================
End Select

Conexao.Close
%>