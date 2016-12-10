<% 
	' Preservar o Idioma
	idioma = Session("idioma")
	' Abandonar a Sessão
	Session.Abandon() 
	' Re-Definir o idioma
	response.write(idioma)
	' Mensagem de expirou
	Session("cliente_msg") = "expirou" 
	' Voltar à página inicial
	If idioma <> "" Then
'		response.Redirect("/?i=" & idioma)
		response.Redirect("/idioma.asp?i=" & idioma)
	Else
		response.Redirect("/")
	End If
%>