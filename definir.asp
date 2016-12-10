<!--#include virtual="/includes/limpar_texto.asp"-->
<%
Session("cliente_idioma") 		= Limpar_Texto(Request("idioma"))
Session("cliente_edicao") 		= Limpar_Texto(Request("edicao"))
Session("cliente_tipo") 		= Limpar_Texto(Request("tipo"))
Session("cliente_formulario") 	= Limpar_Texto(Request("formulario"))

response.Redirect(Limpar_Texto(Request("url")))
%>