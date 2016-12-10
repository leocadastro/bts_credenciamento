<%
If Session("cliente_logado") <> true OR Session("cliente_logado") <> True Then 
	Session("cliente_logado") = false
	%>
    <script language="javascript">
/*
		alert('Sua sessao expirou !\n\nFavor iniciar novamente!');
		document.location = '/';
*/
	</script>
    <%
	Session("cliente_msg") = "expirou"
	response.Redirect("http://cs.btsmedia.biz/sessoes_abandonar.asp")
End If
%>