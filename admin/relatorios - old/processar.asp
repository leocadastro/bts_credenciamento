<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%> 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Credenciamento <%=Year(Now())%> - BTS Informa</title>
<link href="/css/estilos.css" rel="stylesheet" type="text/css">
</head>
<body style="color:#666">
<%
response.Buffer = True
response.Expires = -1
response.AddHeader "Cache-Control", "no-cache"
response.AddHeader "Pragma", "no-cache" 
%>
<!--#include virtual="/admin/inc/gravar_limpar_texto.asp"-->
<!--#include virtual="/admin/inc/acentos_htm.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<!--#include virtual="/admin/scripts/enviar_email.asp"-->
<%
response.Charset = "utf-8" 
response.ContentType = "text/html" 

Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")

'=======================================================================
'	Select 
'		ID_Formulario
'		,Nome
'	FROM Formularios
' **** Resultado
'	ID_Formulario - Nome
'	1 - Empresa
'	2 - Entidades
'	3 - Imprensa
'	4 - Pessoa Física
'	5 - Universidades
'	6 - Alunos

	' Enviar EMAIL
	id_edicao				= limpar_texto(Request("id_edicao"))
	id_idioma				= limpar_texto(Request("id_idioma"))
	ID_Formulario			= limpar_texto(Request("ID_Formulario"))
	Email 					= limpar_texto(Request("Email"))
	Novo_ID_Rel_Cadastro	= limpar_texto(Request("ID"))
	CPF						= limpar_texto(Request("CPF"))
	Nome					= limpar_texto(Request("Nome"))
	Cargo					= limpar_texto(Request("Cargo"))
	Depto					= limpar_texto(Request("Depto"))
	CNPJ					= limpar_texto(Request("CNPJ"))
	Razao					= limpar_texto(Request("Razao"))

'	response.write(id_edicao & "<br>")
'	response.write(id_idioma & "<br>")
'	response.write(ID_Formulario & "<br>")
'	response.write(Email & "<br>")
'	response.write(Novo_ID_Rel_Cadastro & "<br>")
'	response.write(CPF & "<br>")
'	response.write(Nome & "<br>")
'	response.write(Cargo & "<br>")
'	response.write(Depto & "<br>")
'	response.write(CNPJ & "<br>")
'	response.write(Razao & "<br>")


	Enviar_Email id_edicao, id_idioma, ID_Formulario, Email, Novo_ID_Rel_Cadastro, CPF, Nome, Cargo, Depto, CNPJ, Razao

Conexao.Close
%>
</body>
</html>