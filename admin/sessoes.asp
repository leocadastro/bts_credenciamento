<html>
<head>
<meta http-equiv="Content-type" content="text/html; charset=iso-8859-1" />
<title>Administração CSC - Brazil Trade Shows</title>
<link href="/admin/css/bts.css" rel="stylesheet" type="text/css" />
<link href="/admin/css/admin.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table align="center" width="600" border="0" cellspacing="0" cellpadding="0"  class="conteudo_site fs12px">
  <tr>
    <td class="bold">Vari&aacute;veis de Sess&atilde;o - <% =Session.Contents.Count %> Encontradas</td>
    <td align="right" class="bold"><a href="/admin/sessoes_abandonar.asp">Abandonar Sess&otilde;es</a></td>
  </tr>
<%
Dim item, itemloop
For Each item in Session.Contents
  If IsArray(Session(item)) then
    For itemloop = LBound(Session(item)) to UBound(Session(item))
	%>
	  <tr>
		<td width="50%"><% =item %> <% =itemloop %> <font color=blue><% =Session(item)(itemloop) %></font></td>
		<td width="50%"><% =item %> <font color=blue><% =Session.Contents(item) %></font><BR></td>
	  </tr>
	<%
    Next
  Else
  %>
  <tr>
    <td width="50%" align="right"><% =item %></td>
    <td width="50%" style="padding-left:10px;"><font color=blue><% =Session.Contents(item) %></font><BR></td>
  </tr>
  <%
  End If
Next
%>
</table>
</body>
</html>