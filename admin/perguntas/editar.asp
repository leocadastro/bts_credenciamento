<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
id	= Limpar_Texto(Request("id"))

'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")

  SQL_Listar =  "Select " &_
                " ID_Perguntas " &_
                " ,ID_Edicao " &_
                " ,ID_Formulario " &_
                " ,Pergunta_PTB " &_
                " ,Pergunta_ENG " &_
                " ,Pergunta_ESP " &_
                " ,Tipo " &_
                " ,Multiplo " &_
                " ,Ativo " &_
                "From Perguntas " &_
                "Where ID_Perguntas = " & ID

	'response.write("<hr>" & SQL_Listar & "<hr>")

	Set RS_Listar = Server.CreateObject("ADODB.Recordset")
	RS_Listar.Open SQL_Listar, Conexao
	
	If RS_Listar.BOF or RS_Listar.EOF Then
		response.Redirect("default.asp?msg=erro_nao_encontrado")
	Else
		ID_Formulario = RS_Listar("ID_Formulario")
    Pergunta_PTB  = RS_Listar("Pergunta_PTB")
		Pergunta_ENG  = RS_Listar("Pergunta_ENG")
		Pergunta_ESP  = RS_Listar("Pergunta_ESP")
		Tipo          = RS_Listar("Tipo")
    Multiplo      = RS_Listar("Multiplo")
		Ativo 	      = RS_Listar("Ativo")
		RS_Listar.Close
	End If
%>
<html>
<head>
<meta http-equiv="Content-type" content="text/html; charset=iso-8859-1" />
<title>Administração Cred. 2012</title>
<link href="/admin/css/bts.css" rel="stylesheet" type="text/css" />
<link href="/admin/css/admin.css" rel="stylesheet" type="text/css" />
<link href="/css/calendar.css" rel="stylesheet" type="text/css" media="screen">
<script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
<script language="javascript" src="/admin/js/validar_forms.js"></script>
<script language="javascript" src="/js/jquery.maskedinput-1.2.2.min.js"></script>
<script language="javascript" src="/js/colorpicker/colorpicker.js"></script>
<script language="javascript" src="/js/Calendario/calendar.js"></script>
</head>


<script language="javascript">
$(document).ready(function(){
	$('#hora_ini').mask("99:99",{placeholder:"_"});
	$('#aviso').hide();
	<% 
	
	If msg = "" AND Session("admin_msg") <> "" Then msg = Session("admin_msg")
	Select Case msg
		Case "pag_proibida"
			Session("admin_msg") = ""
		%>
		$('#aviso_conteudo').html('Página não permitida !');
		$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000);
		<%
	End Select
	%>
});

function Enviar() {
	var erros = 0;
	$('select:enabled').each(function(i) {
		// Se não for obrigatório
		switch (this.id) {
			default:
				erros += verificar(this.id, false);
				break;
		}
	});
	$('input:enabled').each(function(i) {
		// Se não for obrigatório
		switch (this.id) {
      case 'opcoes_ptb':
        break;
      case 'opcoes_eng':
        break;
      case 'opcoes_esp':
        break; 
      case 'ordem':
        break; 
			default:
				erros += verificar(this.id, false);
				break;
		}
	});
	if (erros == 0) {
		document.cad.submit();	
	} else {
		$('#aviso_conteudo').html('Favor preencher corretamente os campos em destaque.');
		$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000);
	}
}

function Enviar_Opcoes() {
  var erros = 0;
  $('select:enabled').each(function(i) {
    // Se não for obrigatório
    switch (this.id) {
      default:
        erros += verificar(this.id, false);
        break;
    }
  });
  $('input:enabled').each(function(i) {
    // Se não for obrigatório
    switch (this.id) {
      default:
        erros += verificar(this.id, false);
        break;
    }
  });
  if (erros == 0) {
    document.cad_opcoes.submit();  
  } else {
    $('#aviso_conteudo').html('Favor preencher corretamente os campos em destaque.');
    $('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000);
  }
}


function voltar() {
	confirmacao = confirm("Os dados não foram salvos, deseja sair ?")
	if (confirmacao) {
		document.location = 'default.asp';	
	}
}
</script>

<body>
<!--#include virtual="/admin/inc/menu_top.asp"-->
<table width="955" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><img src="/admin/images/img_tabela_branca_top.jpg" width="955" height="15" /></td>
  </tr>
</table>
<table width="955" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td align="center" bgcolor="#FFFFFF"><table width="900" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="100" height="50">&nbsp;</td>
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">Pergunta ID: <%=id%></span></td>
        <td width="100" align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;"><a href="javascript:voltar();"><img src="/admin/images/bt_voltar_admin.gif" width="100" height="48" border="0" /></a></td>
      </tr>
    </table>
      <div id="aviso" style="background-color:#FFFF00; width:100%; text-align:center;" class="t_arial fs11px bold c_preto"> <img src="/admin/images/ico_aviso.gif" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso"> <span id="aviso_conteudo">Aviso</span></div>
      <table width="600" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="14" height="12"><img src="/admin/images/caixa/top_esq.gif" width="14" height="12" /></td>
          <td background="/admin/images/caixa/top_centro.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td width="14" height="12"><img src="/admin/images/caixa/top_dir.gif" width="14" height="12" /></td>
        </tr>
        <tr>
          <td background="/admin/images/caixa/esq.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td><table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
            <form id="cad" name="cad" method="post" action="metodos.asp">
            <input type="hidden" id="acao" name="acao" value="upd_pergunta">
            <input type="hidden" id="id" name="id" value="<%=id%>">
            <input type="hidden" id="id_formulario" name="id_formulario" value="<%=id_formulario%>">
              <tr>
                <td height="30" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">Atualizar</span></td>
                </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Pergunta - PTB</td>
                <td class="titulo_noticias_home"><textarea rows="2" cols="20" name="pergunta_ptb" id="pergunta_ptb" class="admin_txtfield_login" cols="30" rows="3" style="width: 193px; height: 77px; resize:none;"><%=pergunta_ptb%></textarea></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Pergunta - ENG</td>
                <td class="titulo_noticias_home"><textarea rows="2" cols="20" name="pergunta_eng" id="pergunta_eng" class="admin_txtfield_login" cols="30" rows="3" style="width: 193px; height: 77px; resize:none;"><%=pergunta_eng%></textarea></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Pergunta - ESP</td>
                <td class="titulo_noticias_home"><textarea rows="2" cols="20" name="pergunta_esp" id="pergunta_esp" class="admin_txtfield_login" cols="30" rows="3" style="width: 193px; height: 77px; resize:none;"><%=pergunta_esp%></textarea></td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_bts">Tipo</td>
                <td class="titulo_noticias_home">
                  <select id="tipo" name="tipo" class="admin_txtfield_login">
                    <option value="1" <% If Tipo = "1" Then %> selected <% End If %> >Text</option>
                    <option value="2" <% If Tipo = "2" Then %> selected <% End If %> >Radio Button</option>
                    <option value="3" <% If Tipo = "3" Then %> selected <% End If %> >Check Box</option>
                  </select>
                  </td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_bts">Multiplo</td>
                <td class="titulo_noticias_home">
                  <select id="multiplo" name="multiplo" class="admin_txtfield_login">
                    <option value="1" <% If Multiplo = true  Then %> selected <% End If %> >Sim</option>
                    <option value="0" <% If Multiplo = false Then %> selected <% End If %> >Não</option>
                  </select>
                  </td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_tec">&nbsp;</td>
                <td class="titulo_noticias_home">
                  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:left;"></div>
                  <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco cursor" onClick="Enviar();">Atualizar</div>
                  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:right; width:15px; height:17px; float:left;"></div>
                  </td>
                </tr>
            </form>
          </table></td>
          <td background="/admin/images/caixa/dir.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
        </tr>
        <tr>
          <td><img src="/admin/images/caixa/inf_esq.gif" width="14" height="12" /></td>
          <td background="/admin/images/caixa/inf_centro.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td><img src="/admin/images/caixa/inf_dir.gif" width="14" height="12" /></td>
        </tr>
      </table><br/>
      <% 
        If Multiplo = true Then
      %>
      <!-- Cadastrar e lisstar Opções -->
      <table width="600" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="14" height="12"><img src="/admin/images/caixa/top_esq.gif" width="14" height="12" /></td>
          <td background="/admin/images/caixa/top_centro.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td width="14" height="12"><img src="/admin/images/caixa/top_dir.gif" width="14" height="12" /></td>
        </tr>
        <tr>
          <td background="/admin/images/caixa/esq.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td><table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
            <form id="cad_opcoes" name="cad_opcoes" method="post" action="metodos.asp">
            <input type="hidden" id="acao" name="acao" value="add_opcoes">
            <input type="hidden" id="id_pergunta" name="id_pergunta" value="<%=id%>">
              <tr>
                <td height="30" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">Cadastrar Opções</span></td>
                </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Opções - PTB</td>
                <td class="titulo_noticias_home"><input name="opcoes_ptb" id="opcoes_ptb" type="text" class="admin_txtfield_login" size="30" /></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Opções - ENG</td>
                <td class="titulo_noticias_home"><input name="opcoes_eng" id="opcoes_eng" type="text" class="admin_txtfield_login" size="30" /></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Opções - ESP</td>
                <td class="titulo_noticias_home"><input name="opcoes_esp" id="opcoes_esp" type="text" class="admin_txtfield_login" size="30" /></td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_bts">Ordem</td>
                <td class="titulo_noticias_home"><input name="ordem" id="ordem" type="text" class="admin_txtfield_login" size="30" /></td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_bts">Dispon&iacute;vel</td>
                <td class="titulo_noticias_home">
                  <select id="ativo" name="ativo" class="admin_txtfield_login">
                    <option value="0">Não</option>
                    <option value="1">Sim</option>
                  </select>
                  </td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_tec">&nbsp;</td>
                <td class="titulo_noticias_home">
                  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:left;"></div>
                  <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco cursor" onClick="Enviar_Opcoes();">Cadastrar </div>
                  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:right; width:15px; height:17px; float:left;"></div>
                  </td>
                </tr>
            </form>
          </table>
          </td>
          <td background="/admin/images/caixa/dir.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
        </tr>
        <tr>
          <td><img src="/admin/images/caixa/inf_esq.gif" width="14" height="12" /></td>
          <td background="/admin/images/caixa/inf_centro.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td><img src="/admin/images/caixa/inf_dir.gif" width="14" height="12" /></td>
        </tr>
      </table>
    <%
          ' Listar Opções
          SQL_Listar_Opcoes =   "Select " &_
                                " ID_Opcoes " &_
                                " ,ID_Perguntas " &_
                                " ,Opcao_PTB " &_
                                " ,Opcao_ENG " &_
                                " ,Opcao_ESP " &_
                                " ,Ordem " &_
                                " ,Ativo " &_
                                " ,Data_Cadastro " &_
                                "From Perguntas_Opcoes " &_
                                "Where ID_Perguntas = " & ID & " " &_
                                "Order By Ordem ASC"

          'response.write("<hr>" & SQL_Listar & "<hr>")
          Set RS_Listar_Opcoes = Server.CreateObject("ADODB.Recordset")
          RS_Listar_Opcoes.Open SQL_Listar_Opcoes, Conexao
    %>
      <br>
         <table width="900" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="14" height="12"><img src="/admin/images/caixa/top_esq.gif" width="14" height="12"></td>
          <td background="/admin/images/caixa/top_centro.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12"></td>
          <td width="14" height="12"><img src="/admin/images/caixa/top_dir.gif" width="14" height="12"></td>
        </tr>
        <tr>
          <td background="/admin/images/caixa/esq.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12"></td>
          <td>
      <!--#include virtual="/admin/inc/paginacao.asp"-->
            <%  If RS_Listar_Opcoes.BOF or RS_Listar_Opcoes.EOF Then  %>
            <p align="center" class="titulo_menu_site_carne">N&atilde;o foi encontrado nenhum registro</p>
            <% End If %>
            <%
        Contador = 0
        If not RS_Listar_Opcoes.BOF or not RS_Listar_Opcoes.EOF Then
    %>
            <table width="100%" border="0" cellpadding="2" cellspacing="1" class="fs11px t_arial">
              <tr>
                <td width="50" align="center" class="borda_dir linha_16px"><b>ID</b></td>
                <td class="borda_dir linha_16px" align="center"><b>Portugu&ecirc;s</b></td>
                <td class="borda_dir linha_16px" align="center"><b>Ingl&ecirc;s</b></td>
                <td class="borda_dir linha_16px" align="center"><b>Espanhol</b></td>
                <td width="40" align="center" class="borda_dir linha_16px"><b>Ordem</b></td>
                <td width="65" align="center" class="borda_dir linha_16px"><b>Disponível</b></td>
                <td width="100" align="center" class="borda_dir linha_16px"><b>Data</b></td>
                <td width="40" align="center" class="linha_16px"><b>Editar</b></td>
              </tr>
              <%
            RS_Listar_Opcoes.MoveFirst
            While Not RS_Listar_Opcoes.EOF
            %>
              <tr bgcolor="#FFFFFF" 
                onMouseOver="$(this).attr('bgcolor','#FFFF00'); " 
                onMouseOut="$(this).attr('bgcolor','#FFFFFF'); "  >
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'editar_opcoes.asp?id=<%=RS_Listar_Opcoes("ID_Opcoes")%>&id_pergunta=<%=ID%>';" align="left"><%=RS_Listar_Opcoes("ID_Opcoes")%></td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'editar_opcoes.asp?id=<%=RS_Listar_Opcoes("ID_Opcoes")%>&id_pergunta=<%=ID%>';" align="left"><%=RS_Listar_Opcoes("Opcao_PTB")%></td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'editar_opcoes.asp?id=<%=RS_Listar_Opcoes("ID_Opcoes")%>&id_pergunta=<%=ID%>';" align="left"><%=RS_Listar_Opcoes("Opcao_ENG")%></td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'editar_opcoes.asp?id=<%=RS_Listar_Opcoes("ID_Opcoes")%>&id_pergunta=<%=ID%>';" align="left"><%=RS_Listar_Opcoes("Opcao_ESP")%></td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'editar_opcoes.asp?id=<%=RS_Listar_Opcoes("ID_Opcoes")%>&id_pergunta=<%=ID%>';" align="center"><%=RS_Listar_Opcoes("Ordem")%></td>
                <td class="borda_dir linha_16px " align="center">
                <%
          ativo = RS_Listar_Opcoes("Ativo")
          If ativo = true OR ativo = 1 Then
            %><img src="/img/geral/icones/ok.gif" width="20" height="20" alt="ativo" title="ativo" onClick="document.location = 'metodos.asp?id=<%=id%>&id_op=<%=RS_Listar_Opcoes("ID_Opcoes")%>&acao=desativar_opcoes';"><%
          Else
            %><img src="/img/geral/icones/nok.gif" width="20" height="20" alt="desativado" title="desativado" onClick="document.location = 'metodos.asp?id=<%=id%>&id_op=<%=RS_Listar_Opcoes("ID_Opcoes")%>&acao=ativar_opcoes';"><%
                  End If
        %>
                </td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'editar_opcoes.asp?id=<%=RS_Listar_Opcoes("ID_Opcoes")%>&id_pergunta=<%=ID%>';" align="left"><%=FormatDateTIme(RS_Listar_Opcoes("Data_Cadastro"),2)%></td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'editar_opcoes.asp?id=<%=RS_Listar_Opcoes("ID_Opcoes")%>&id_pergunta=<%=ID%>';" align="center"><img src="/admin/images/ico_pg_prox.gif" width="15" height="15" alt="Editar Pergunta" title="Editar Pergunta" border="0"></td>
              </tr>
              <%
                RS_Listar_Opcoes.MoveNext()
            Wend
            RS_Listar_Opcoes.Close
      %>
            </table>
            <%
        End If
        %>

    <%
      End If
    %>
    </td>
          <td background="/admin/images/caixa/dir.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12"></td>
        </tr>
        <tr>
          <td><img src="/admin/images/caixa/inf_esq.gif" width="14" height="12"></td>
          <td background="/admin/images/caixa/inf_centro.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12"></td>
          <td><img src="/admin/images/caixa/inf_dir.gif" width="14" height="12"></td>
        </tr>
      </table>
    <!-- Listar as Opções Cadastradas -->
    </td>
  </tr>
</table>
<table width="955" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="/admin/images/img_tabela_branca_inferior.jpg" width="955" height="15" /></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>