<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="33%" align="center">&nbsp;</td>
    <td width="870" align="center"><!-- Cabecalho -->
      <table width="870" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td height="30">&nbsp;</td>
        </tr>
        <tr>
          <td height="104"><table width="870" height="104" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="320" rowspan="2"><a href="/admin/"><img src="/img/geral/cabecalho_logo.gif" width="188" height="103" border="0" /></a></td>
              <td height="74" valign="bottom"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td valign="bottom"><a href="/admin/menu.asp"><img src="/admin/images/bt_menu_admin.gif" width="76" height="30" border="0" /></a></td>
                      <td><div align="right">
                        <%if Session("admin_id_perfil") = "1" Then ' Admin %>
                        <img src="/admin/images/menu_top_admin.gif" width="350" height="36" border="0" usemap="#MapAdminMap" />
                        <%else%>
                        <img src="/admin/images/menu_top.gif" width="350" height="36" border="0" usemap="#MapMap" />
                        <%end if%>
                      </div></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td height="25" colspan="2" align="right" class="menu_bts_padrao"><%=Session("admin_txt_nome")%> - <%=Session("admin_txt_email")%></td>
                </tr>
              </table>
                <map name="MapMap" id="MapMap">
                  <area shape="rect" coords="284,8,330,27" href="/admin/logoff.asp" />
                  <area shape="rect" coords="170,8,245,28" href="/admin/administradores/atualizar.asp" />
                </map>
                <map name="MapAdminMap" id="MapAdminMap">
                  <area shape="rect" coords="33,8,130,27" href="/admin/administradores/" />
                  <area shape="rect" coords="170,8,245,28" href="/admin/administradores/atualizar.asp" />
                  <area shape="rect" coords="284,8,330,27" href="/admin/logoff.asp" />
                </map></td>
            </tr>
            <tr>
              <td height="30" background="/img/geral/fundo_cabecalho.gif"><img src="/admin/images/faixa_administracao.gif" width="169" height="28" hspace="1" /><img src="/img/geral/nome_cabecalho.gif" width="199" height="28" /></td>
            </tr>
          </table></td>
        </tr>
      </table>
      <!-- Cabecalho --></td>
    <td width="33%" align="center" valign="top">
    <!-- Faixa Cabecalho -->
      <div style="background:url(/img/geral/fundo_cabecalho.gif); height:30px; width:100%; margin-top:104px;"></div>
    <!-- Faixa Cabecalho	 -->
    </td>
  </tr>
</table>