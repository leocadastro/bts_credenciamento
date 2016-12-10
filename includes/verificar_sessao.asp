<%
If Session("cliente_idioma") = "" Then 
	%>
    <script language="javascript">
		alert('Sua sessao expirou !\n\nFavor iniciar novamente!');
		document.location = '/';
	</script>
    <%
End IF
%>