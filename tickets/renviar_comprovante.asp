<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%> 
<%
Response.Expires = -1
Response.AddHeader "Cache-Control", "no-cache"
Response.AddHeader "Pragma", "no-cache" 
%>
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/acentos_htm.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<!--#include virtual="/scripts/enviar_email_senha.asp"-->
<%
response.Charset = "utf-8" 
response.ContentType = "text/html" 

'=======================================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'=======================================================================

' Pegando os Parâmetros necessários
'=======================================================================
CodPedido		= Limpar_Texto(Request("pedido"))
ID_Visitante	= Limpar_Texto(Request("visitante"))
ID_Edicao 		= Limpar_Texto(Request("edicao"))
Idioma 			= Limpar_Texto(Request("idioma"))

'=======================================================================
' Verificando valor do campo documento
'=======================================================================
If Len(ID_Visitante) <> 0 Or Len(CodPedido) <> 0 Then
	'=======================================================================
	' Verificando cadastro no banco novo Credenciamento_2012
	'=======================================================================
	
	'==================================================
	' Nova Implementacao - 14/05/2014 - Leandro Santiago | HD 8183
	' INICIO - Verificar Edicao do Visitante para que nao envie e-mail para Visitante errado
	'==================================================
	
	SQL_Recuperar_Senha =	"Select " &_
							"	Top 1 " &_
							"	V.ID_Visitante, " &_
							"	V.Nome_Completo, " &_
							"	V.Email, " &_
							"	V.Senha " &_
							"From Visitantes as V " &_
							"	Inner Join Relacionamento_Cadastro as RC " &_
							"		On RC.ID_Visitante = V.ID_Visitante " &_
							"		And RC.ID_Edicao = " & ID_Edicao & " " &_
							"Where " &_
							"	V.ID_Visitante = '" & ID_Visitante & "'  " &_
							"Order by V.ID_Visitante DESC"
	'==================================================
	' FIM - Verificar Edicao do Visitante para que nao envie e-mail para Visitante errado
	'==================================================
	
	Set RS_Recuperar_Senha = Server.CreateObject("ADODB.Recordset")
	RS_Recuperar_Senha.Open SQL_Recuperar_Senha, Conexao, 3, 3
	
	If Not RS_Recuperar_Senha.BOF or Not RS_Recuperar_Senha.EOF Then
	 
		ID_Visitante	= RS_Recuperar_Senha("ID_Visitante")
		Email 			= RS_Recuperar_Senha("Email")
		Nome 			= RS_Recuperar_Senha("Nome_Completo")
		Senha 			= Trim(RS_Recuperar_Senha("Senha"))

		'response.write ("senha: " & senha & "<br>")
		'response.write ("nova_senha: " & nova_senha & "<br>")
		
		Enviar_Email_Senha ID_Edicao, Idioma, "", "", Email, Nome, Senha, "Enviar_Ticket", CodPedido, ID_Visitante
		
		%>{ 'retorno' : 'email enviado', 'email' : '<%=Lcase(email)%>', 'nome' : '<%=nome%>' }<%
		

	End If
Else 
	%>{ 'retorno' : 'login invalido' }<%
End If

Conexao.Close
%>