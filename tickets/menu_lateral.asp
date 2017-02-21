<script language="javascript">
function link(qual){
	setTimeout(function() {
		$('#conteudo').css( {"z-index": 10 }).hide("slide", { direction: "left" }, 1000);
		var urls = '/tickets/' + qual;
		document.location = urls;
	},1000);
}
</script>

<%
texto_menu = ""

SQL_Pedidos_Menu = 	"Select " &_
					"	* " &_
					"From Pedidos " &_
					"Where " &_
					"	ID_Rel_Cadastro = " & Session("cliente_cadastro") & " " &_
					"	And ID_Visitante = " & Session("cliente_visitante") & " " &_
					"	And Status_Pedido = 1 " &_
					"	And ID_Edicao = " & Session("Cliente_Edicao") & " "
Set RS_Pedidos_Menu = Server.CreateObject("ADODB.Recordset")
RS_Pedidos_Menu.Open SQL_Pedidos_Menu, Conexao, 3, 3

If Not RS_Pedidos_Menu.Eof Then
	texto_menu 	= "Continuar Pedido"
	link_menu	= "novo_pedido.asp"
Else
	texto_menu 	= "Novo Pedido"
	link_menu	= "termo.asp"
End If

' Se estiver na Página de RETORNO da TRANSAÇÃO exibir NOVA COMPRA
If Request.ServerVariables("URL") = "/tickets/retorno_exibir.asp" Then
	' E meu pedido nao foi aprovado
	If Aprovacao = "True"  Then
		texto_menu 	= "Novo Pedido"
		link_menu	= "termo.asp"
	End If
ElseIf Request.ServerVariables("URL") = "/tickets/pagamento.asp" Then
	texto_menu 	= "Alterar seu Pedido"
	link_menu	= "novo_pedido.asp"
End If
%>
<div style="float: left; width: 210px;">
    <table width="210" border="0" cellpadding="0" cellspacing="0">
        <tbody>
        <tr>
            <td height="26" bgcolor="#414042" class="arial fs_13px cor_branco b" style="padding-left:10px;">Menu</td>
        </tr>
        <tr>
            <td align="center">
            <table width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom:10px;">
                <tbody>
                <%
				SQL_Pedidos = 	"Select " &_
								"	* " &_
								"From Pedidos As P " &_
								"Inner Join Pedidos_Carrinho As C " &_
								"	On C.ID_Pedido = P.ID_Pedido " &_
								"Where " &_
								"	P.ID_Edicao = " & Session("cliente_edicao") & " " &_
								"	And C.ID_Visitante = " & Session("cliente_visitante") & " " &_
								"	And P.Status_Pedido <> 1 "
								'"	And Status_Pedido <> 3"
				'Response.Write(SQL_Pedidos)
				Set RS_Pedidos = Server.CreateObject("ADODB.Recordset")
				RS_Pedidos.Open SQL_Pedidos, Conexao, 3, 3

				If Not RS_Pedidos.Eof Then
				%>
                <tr class="cursor menu" onmouseover="menu(this, 'over');" onmouseout="menu(this, 'out');" onclick="link('status.asp');" style="background-color: rgb(255, 255, 255);">
                    <td height="30" style="padding-left:10px;"><span class="arial fs_12px cor_cinza1"><img src="/img/geral/icones/item_menu.gif" width="20" height="10">Meus Pedidos</span></td>
                </tr>
                <tr>
                    <td height="4"><img src="/img/geral/spacer.gif" width="1" height="4"></td>
                </tr>
                <%End If%>
                <tr class="cursor menu" onmouseover="menu(this, 'over');" onmouseout="menu(this, 'out');" onclick="link('<%=link_menu%>');" style="background-color: rgb(255, 255, 255);">
                    <td height="30" style="padding-left:10px;">
                    <span class="arial fs_12px cor_cinza1">
                    	<img src="/img/geral/icones/item_menu.gif" width="20" height="10"><%=texto_menu%>
                    </span></td>
                </tr>
                <tr>
                    <td height="4"><img src="/img/geral/spacer.gif" width="1" height="4"></td>
                </tr>
                <%  '//Apenas a ABF
                    If Session("cliente_edicao") = "56" Then
                 %>
                <tr class="cursor menu" onmouseover="menu(this, 'over');" onmouseout="menu(this, 'out');" style="background-color: rgb(255, 255, 255);">
                    <td height="30" style="padding-left:10px;">
                    <span class="arial fs_12px cor_cinza1">
                    	<img src="/img/geral/icones/item_menu.gif" width="20" height="10">Suporte: visitante.abf@informa.com
                    </span></td>
                </tr>
                 <% Else %>
                <!--tr class="cursor menu" onmouseover="menu(this, 'over');" onmouseout="menu(this, 'out');" style="background-color: rgb(255, 255, 255);">
                    <td height="30" style="padding-left:10px;">
                    <a href="http://www.easychat.com.br/easy/iframe_w33.php?chat_id=2103&amp;clie_id=2039&amp;check_sum=6181" target="_blank" style="text-decoration: none">
                    <span class="arial fs_12px cor_cinza1">
                    	<img src="/img/geral/icones/item_menu.gif" width="20" height="10">Suporte: Chat Online
                    </span>
                    </a>
                    </td>
                </tr>
                <tr>
                    <td height="4"><img src="/img/geral/spacer.gif" width="1" height="4"></td>
                </tr-->
                <tr class="cursor menu" onmouseover="menu(this, 'over');" onmouseout="menu(this, 'out');" style="background-color: rgb(255, 255, 255);">
                    <td height="30" style="padding-left:10px;">
                    <span class="arial fs_12px cor_cinza1">
                    	<img src="/img/geral/icones/item_menu.gif" width="20" height="10">Suporte: (11) 3598-7834 <br/>ou <a href="mailto:visitante.abf@informa.com">visitante.abf@informa.com</a>
                    </span></td>
                </tr>
                <%  End If %>
				<tr><td>&nbsp;</td>
				</tr>
				<tr>
				<td align="center">
				<!-- PayPal Logo --><img  src="https://www.paypal-brasil.com.br/logocenter/util/img/compra_segura_horizontal.png" border="0" alt="Imagens de solução" /><!-- PayPal Logo -->
				</td>
				</tr>
                </tbody>
            </table>
            </td>
        </tr>
        </tbody>
    </table>
</div>
