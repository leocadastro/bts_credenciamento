<!--#include virtual="/includes/limpar_texto.asp"-->
<%
idioma = Limpar_Texto(Request("i"))
If idioma = "" Or Cint(idioma) < 1 Or Cint(idioma) > 3 Then 
	Session("cliente_idioma") = 1 ' Portugues
Else
	Session("cliente_idioma") = idioma
End If
%>