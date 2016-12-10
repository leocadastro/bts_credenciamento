<% If SQL_Textos <> "" Then %>
	<table style="border:1px solid #000;" cellspacing="3" cellpadding="3" class="verdana fs_10px cor_preto">
	<tr>
	<td style="border-right:1px solid #000; border-bottom:1px solid #000;">ID</td>
	<td style="border-right:1px solid #000; border-bottom:1px solid #000;">Ident</td>
	<td style="border-right:1px solid #000; border-bottom:1px solid #000;">Texto / C&oacute;digo</td>
	<td style="border-right:1px solid #000; border-bottom:1px solid #000;">IMG / C&oacute;digo</td>
	</tr>
	<% 
	x = 0
	For i = Lbound(textos_array) to Ubound(textos_array)
		If x = 1 Then
			bg = "#ffffff"
			x = 0
		Else 
			bg = "#FFFFCC"
			x = 1
		End If
	%>
        <tr bgcolor="<%=bg%>">
        <td style="border-right:1px solid #000; border-bottom:1px solid #000;"><b><%=i%></b></td>
        <td style="border-right:1px solid #000; border-bottom:1px solid #000;">
			<%=textos_array(i)(1)%><!--<br size="1"><b>&lt;%=textos_array(<%=i%>)(1)%&gt;</b>-->
        </td>
        <td style="border-right:1px solid #000; border-bottom:1px solid #000;">
        <% If textos_array(i)(2) <> "" Then %>
			<%=textos_array(i)(2)%><hr size="1"><b>&lt;%=textos_array(<%=i%>)(2)%&gt;</b>
        <% End If %>
        </td>
        <td style="border-right:1px solid #000; border-bottom:1px solid #000;">
        <% If textos_array(i)(3) <> "" Then %>
            <img src="<%=textos_array(i)(3)%>"><hr size="1"><b>&lt;%=textos_array(<%=i%>)(3)%&gt;</b>
        <% End If %>
        &nbsp;
        </td>
        </tr>
	<%Next%>
	</table>
    <hr />
<% End If %>
    <% If SQL_Textos_Produto <> "" Then %>
    	<big><b>Produto</b></big><br />
        <table style="border:1px solid #000;" cellspacing="3" cellpadding="3" class="verdana fs_10px cor_preto">
        <tr>
        <td style="border-right:1px solid #000; border-bottom:1px solid #000;">ID</td>
        <td style="border-right:1px solid #000; border-bottom:1px solid #000;">Ident</td>
        <td style="border-right:1px solid #000; border-bottom:1px solid #000;">Texto / C&oacute;digo</td>
        <td style="border-right:1px solid #000; border-bottom:1px solid #000;">IMG / C&oacute;digo</td>
        </tr>
        <% 
        x = 0
        For i = Lbound(produto_textos_array) to Ubound(produto_textos_array)
            If x = 1 Then
                bg = "#ffffff"
                x = 0
            Else 
                bg = "#FFFFCC"
                x = 1
            End If
        %>
            <tr bgcolor="<%=bg%>">
            <td style="border-right:1px solid #000; border-bottom:1px solid #000;"><b><%=i%></b></td>
            <td style="border-right:1px solid #000; border-bottom:1px solid #000;">
                <%=produto_textos_array(i)(1)%><!--<br size="1"><b>&lt;%=textos_array(<%=i%>)(1)%&gt;</b>-->
            </td>
            <td style="border-right:1px solid #000; border-bottom:1px solid #000;">
            <% If produto_textos_array(i)(2) <> "" Then %>
                <%=produto_textos_array(i)(2)%><hr size="1"><b>&lt;%=produto_textos_array(<%=i%>)(2)%&gt;</b>
            <% End If %>
            </td>
            <td style="border-right:1px solid #000; border-bottom:1px solid #000;">
            <% If produto_textos_array(i)(3) <> "" Then %>
                <img src="<%=produto_textos_array(i)(3)%>"><hr size="1"><b>&lt;%=produto_textos_array(<%=i%>)(3)%&gt;</b>
            <% End If %>
            &nbsp;
            </td>
            </tr>
        <%Next%>
        </table>
        <hr />
    <% End If %>