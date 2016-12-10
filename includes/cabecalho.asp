<%
	Select Case (Session("cliente_idioma"))
		Case "1"
			bt_ajuda	= "/img/botoes/ajuda.gif"
			titulo 		= "/img/geral/nome_cabecalho.gif"
		Case "2"
			bt_ajuda 	= "/img/botoes/ajuda_esp.gif"
			titulo 		= "/img/geral/nome_cabecalho_esp.gif"
		Case "3"
			bt_ajuda 	= "/img/botoes/ajuda_eng.gif"
			titulo 		= "/img/geral/nome_cabecalho_eng.gif"
		Case Else
			bt_ajuda 	= "/img/botoes/ajuda.gif"
			titulo 		= "/img/geral/nome_cabecalho.gif"
	End Select
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="33%" align="center">&nbsp;</td>
    <td width="870" align="center">
    <!-- Cabecalho -->
        <table width="870" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr>
            <td height="30">&nbsp;</td>
          </tr>
          <tr>
            <td height="104">
                <table width="870" height="104" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="320" rowspan="2"><a href="javascript:sair();"><img src="/img/geral/informa_exhibition.png" width="188" height="103" border="0"></a></td>
                    <td height="74" valign="bottom"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                          <td width="80" height="32"><!--a href="javascript:sair();"><img src="/img/botoes/home.gif" border="0"></a-->&nbsp;</td>
                          <td width="80" height="32">&nbsp;
                          <!--<a href="javascript:jAlert('1','2');"><img id="img_bt_ajuda" src="<%=bt_ajuda%>" title="" alt="" border="0">--></a>
                          </td>
                          <td align="right">
                          <% If Session("cliente_idioma") = "" Then %>
                               <img src="/img/geral/idiomas.gif" hspace="20" border="0" usemap="#idiomas" />
                          <% End If %>
                          </td>
                      </tr>
                    </table></td>
                  </tr>
                  <tr>
                    <td height="30" background="/img/geral/fundo_cabecalho.gif"><img id="img_tit_cabecalho" src="<%=titulo%>" title="" alt=""></td>
                  </tr>
                </table>
            </td>
          </tr>
      </table>
    <!-- Cabecalho -->
    </td>
    <td width="33%" align="center" valign="top">
    <!-- Faixa Cabecalho -->
    	<div style="background:url(/img/geral/fundo_cabecalho.gif); height:30px; width:100%; margin-top:104px;"></div>
    <!-- Faixa Cabecalho	 -->
    </td>
  </tr>
</table>
<map name="idiomas">
  <area shape="rect" coords="0,0,45,28" 	href="javascript:idioma(1);" alt="Portugu&ecirc;s" title="Portugu&ecirc;s" />
  <area shape="rect" coords="48,0,93,28" 	href="javascript:idioma(2);" alt="Espa&ntilde;hol" title="Espa&ntilde;hol" />
  <area shape="rect" coords="95,0,138,28" 	href="javascript:idioma(3);" alt="English" title="English" />
</map>