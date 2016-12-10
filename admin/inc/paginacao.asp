<% If TotalPaginas > 1 Then %>
<% 
' Deixar a página atual sempre no meio se possível ( intpage = Pagina Atual )
If (intpage - 5) >= 0 Then 
	inicio = intpage - 4
Else
	inicio = 1
End If

' Final sempre de 10 paginas
If intpage <= 5 Then
	fim = 10
Else 
	fim = intpage + 4
End If

' Final nao pode ser maior que TotalPaginas
If fim > TotalPaginas Then fim = TotalPaginas

' Correcao do inicio
If fim - intpage <= 3 then 
	dif = 4 -(TotalPaginas - intpage)
	inicio = inicio - dif
	If inicio <= 0 then inicio = 1
End If
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="35" align="center">
    <table border="0" cellspacing="0" cellpadding="0">
      <tr>
      	<% If intpage > 1 Then %>
        	<td width="25" height="25" align="center" style="cursor:pointer;" onclick="document.location='<%=strURL%>&acp=a&pag=<%=intpage%>';"><img src="/admin/images/ico_pg_voltar.gif" alt="Pág. Anterior" width="24" height="24" /></td>
        <% Else %>
        	<td width="25" height="25" >&nbsp;</td>
        <% End If %>
        <% If inicio > 1 Then %>
        	<td width="25" height="25" align="center"><div class="icone_paginacao" style="cursor:default;">...</div></td>
        <% End If %>
        
        <% For i = inicio to fim %>
        	<% If Cint(intpage) = Cint(i) Then ico = "icone_pagina_atual" Else ico = "icone_paginacao" End If %>
	        <td width="25" height="25" align="center"><div class="<%=ico%>" onclick="document.location='<%=strURL%>&acp=n&pag=<%=i%>';"><%=i%></div></td>
        <% Next %>
        
        <% If fim < TotalPaginas Then %>
        	<td width="25" height="25" align="center"><div class="icone_paginacao" style="cursor:default;">...</div></td>
        <% ElseIf TotalPaginas >= 10 Then %>
	        <td width="25" height="25" align="center">&nbsp;</td>
        <% End If %>
        <% If Cint(intpage) < Cint(fim) Then %>
        	<td width="25" height="25" align="center" style="cursor:pointer;" onclick="document.location='<%=strURL%>&acp=p&pag=<%=intpage%>';"><img src="/admin/images/ico_pg_prox.gif" alt="Próxima Pág." width="24" height="24" /></td>
		<% Else %>
        	<td width="25" height="25" >&nbsp;</td>
        <% End If %> 
      </tr>
    </table>
    </td>
  </tr>
</table>
<% 	End IF %>