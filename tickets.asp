<!--#include virtual="/includes/limpar_texto.asp"-->
<%
Session("cliente_idioma") 		= 1
Session("cliente_edicao") 		= 46
Session("cliente_tipo") 		= 10
Session("cliente_formulario") 	= 4

URL = "/tickets/"

response.Redirect(url)
%>