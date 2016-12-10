<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
ord	= Limpar_Texto(Request("ord"))
idp	= Limpar_Texto(Request("idp"))

'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
	
	SQL_Eventos = "Select " &_
					"	* " &_
					"From Eventos_Edicoes AS Ee " &_
					"	Inner Join Eventos AS E" &_
					"	ON Ee.ID_Evento = E.ID_Evento Order by E.ID_EVENTO DESC"
	Set RS_Eventos = Server.CreateObject("ADODB.Recordset")
	RS_Eventos.Open SQL_Eventos, Conexao
	
	SQL_Tipo = 		"Select " &_
					"	ID_Tipo, " &_
					"	Tipo " &_
					"From Paginas_Tipo "
	Set RS_Tipo = Server.CreateObject("ADODB.Recordset")
	RS_Tipo.Open SQL_Tipo, Conexao
	
	SQL_Listar = 	"Select " &_
					"	ID_Texto, " &_
					"	ID_Idioma, " &_
					"	ID_Pagina, " &_
					"	ID_Tipo, " &_
					"	Ordem, " &_
					"	Identificacao, " &_
					"	Texto, " &_
					"	URL_Imagem " &_
					"From Paginas_Textos " &_
					"Where " &_
					"	Ordem = " & ord & " " &_
					"	AND ID_Pagina = " & idp & " " &_
					"Order by Ordem, ID_Idioma"

'	response.write("<hr>" & SQL_Listar & "<hr>")

	Set RS_Listar = Server.CreateObject("ADODB.Recordset")
	RS_Listar.Open SQL_Listar, Conexao
	
	' Se houver registros
	If not RS_Listar.BOF or not RS_Listar.EOF Then
		' Lopping
		While not RS_Listar.EOF 
			If RS_Listar("id_idioma") = "1" Then
				id_pagina 		= RS_Listar("id_pagina")
				id_tipo 		= RS_Listar("id_tipo")
				identificacao 	= RS_Listar("identificacao")
	
				id_texto_ptb 	= RS_Listar("id_texto")
				texto_ptb		= RS_Listar("texto")
				img_ptb			= RS_Listar("url_imagem")
			End If
			If RS_Listar("id_idioma") = "3" Then
				id_pagina 		= RS_Listar("id_pagina")
				id_tipo 		= RS_Listar("id_tipo")
				If identificacao = "" Then identificacao 	= RS_Listar("identificacao")
	
				id_texto_eng 	= RS_Listar("id_texto")
				texto_eng		= RS_Listar("texto")
				img_eng			= RS_Listar("url_imagem")
			End If
			If RS_Listar("id_idioma") = "2" Then
				id_pagina 		= RS_Listar("id_pagina")
				id_tipo 		= RS_Listar("id_tipo")
				If identificacao = "" Then identificacao 	= RS_Listar("identificacao")
	
				id_texto_esp 	= RS_Listar("id_texto")
				texto_esp		= RS_Listar("texto")
				img_esp			= RS_Listar("url_imagem")
			End If
			RS_Listar.MoveNext
		Wend
	
		RS_Listar.Close
	End If

%>
<html>
<head>
<meta http-equiv="Content-type" content="text/html; charset=iso-8859-1" />
<title>Administração CSC - BTS Informa</title>
<link href="/admin/css/bts.css" rel="stylesheet" type="text/css" />
<link href="/admin/css/admin.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
<script language="javascript" src="/admin/js/validar_forms.js"></script>
</head>


<script language="javascript">
$(document).ready(function(){
	$('#aviso').hide();
	$('#protheus').select().focus();
	<% 
	msg = Request("msg")
	Select Case msg
		Case "pag_proibida"
			Session("admin_msg") = ""
			%>
			$('#aviso_conteudo').html('Página não permitida !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000);
			<%
		Case "add_ok"
			%>
			$('#aviso_conteudo').html('Evento adicionado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
			<%
		Case "add_erro_existe"
			%>
			$('#aviso_conteudo').html('Erro - Evento informado já existe !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
			<%
		Case "upd_ok"
			%>
			$('#aviso_conteudo').html('Evento atualizado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
			<%
		Case "atv_ok"
			%>
			$('#aviso_conteudo').html('Evento ativado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
			<%
		Case "des_ok"
			%>
			$('#aviso_conteudo').html('Evento desativado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
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
			case 'id_texto_ptb':
				break;
			case 'id_texto_eng':
				break;
			case 'id_texto_esp':
				break;
			case 'texto_ptb':
				break;
			case 'texto_eng':
				break;
			case 'texto_esp':
				break;
			case 'img_ptb':
				break;
			case 'img_eng':
				break;
			case 'img_esp':
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
		$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
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
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">P&aacute;gina ID: <%=ID_Pagina%></span></td>
        <td width="100" align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;"><a href="javascript:history.back();"><img src="/admin/images/bt_voltar_admin.gif" width="100" height="48" border="0" /></a><a href="default.asp"></a></td>
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
              <input type="hidden" id="acao" name="acao" value="upd_texto">
              <input type="hidden" id="id_pagina" name="id_pagina" value="<%=id_pagina%>">
              <input type="hidden" id="id_texto_ptb" name="id_texto_ptb" value="<%=id_texto_ptb%>">
              <input type="hidden" id="id_texto_eng" name="id_texto_eng" value="<%=id_texto_eng%>">
              <input type="hidden" id="id_texto_esp" name="id_texto_esp" value="<%=id_texto_esp%>">
              <tr>
                <td height="30" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">Atualizar Texto </span></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts"> Evento</td>
                <td class="titulo_noticias_home"><select id="id_edicao" name="id_edicao" class="admin_txtfield_login">
                  <option value="-">-- Selecione --</option>
                  <%
				If not RS_Eventos.BOF or not RS_Eventos.EOF Then
				%>
                  <option value="Null"
                  <% If isNull(id_edicao) OR Len(Trim(id_edicao)) = 0 Then response.Write("selected") End If %>
                  >Geral</option>
                <%
					While not RS_Eventos.EOF
					
					selecionado = ""
					If Cstr(id_edicao) = Cstr(RS_Eventos("ID_Edicao")) Then selecionado = " selected "
					%>
                  <option value="<%=RS_Eventos("ID_Edicao")%>"><%=RS_Eventos("ID_Edicao")%> - <%=RS_Eventos("Ano")%> - <%=RS_Eventos("Nome_PTB")%></option>
                  <%
						RS_Eventos.MoveNext
					Wend
					RS_Eventos.Close
				End If
				%>
                </select></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts"> Tipo de Tradu&ccedil;&atilde;o:</td>
                <td class="titulo_noticias_home"><select id="id_tipo" name="id_tipo" class="admin_txtfield_login">
                  <option value="-">Selecione</option>
					<%
                    If not RS_Tipo.BOF or not RS_Tipo.EOF Then
						While not RS_Tipo.EOF
							selecionado = ""
							If Cstr(id_tipo) = Cstr(RS_Tipo("id_tipo")) Then selecionado = " selected "
							%>
							<option value="<%=RS_Tipo("ID_Tipo")%>" <%=selecionado%>><%=RS_Tipo("Tipo")%></option>
							<%
							RS_Tipo.MoveNext
						Wend
						RS_Tipo.Close
                    End If
                    %>
                  </select></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Ordem:</td>
                <td class="titulo_noticias_home"><%=ord%><input type="hidden" id="ordem" name="ordem" value="<%=ord%>"></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Texto Identifica&ccedil;&atilde;o</td>
                <td class="titulo_noticias_home"><input name="identificacao" id="identificacao" type="text" class="admin_txtfield_login" size="30" value="<%=Identificacao%>" /></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Texto PTB</td>
                <td class="titulo_noticias_home"><textarea name="texto_ptb" cols="30" rows="3" class="admin_txtfield_login" id="texto_ptb"><%=texto_ptb%></textarea></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Texto ENG</td>
                <td class="titulo_noticias_home"><textarea name="texto_eng" cols="30" rows="3" class="admin_txtfield_login" id="texto_eng"><%=texto_eng%></textarea></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Texto ESP</td>
                <td class="titulo_noticias_home"><textarea name="texto_esp" cols="30" rows="3" class="admin_txtfield_login" id="texto_esp"><%=texto_esp%></textarea></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Url Imagem PTB</td>
                <td class="titulo_noticias_home"><input name="img_ptb" id="img_ptb" type="text" class="admin_txtfield_login" size="30"  value="<%=img_ptb%>" />
                  <img src="/admin/images/ico_preview_20_b.gif" alt="Visualizar Imagem" title="Visualizar Imagem" class="cursor" onClick="window.open($('#logo_faixa').val());" width="20" height="20" align="absmiddle"></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Url Imagem ENG</td>
                <td class="titulo_noticias_home"><input name="img_eng" id="img_eng" type="text" class="admin_txtfield_login" size="30"  value="<%=img_eng%>" />
                  <img src="/admin/images/ico_preview_20_b.gif" alt="Visualizar Imagem" title="Visualizar Imagem" class="cursor" onClick="window.open($('#logo_faixa').val());" width="20" height="20" align="absmiddle"></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Url Imagem ESP</td>
                <td class="titulo_noticias_home"><input name="img_esp" id="img_esp" type="text" class="admin_txtfield_login" size="30"  value="<%=img_esp%>" />
                  <img src="/admin/images/ico_preview_20_b.gif" alt="Visualizar Imagem" title="Visualizar Imagem" class="cursor" onClick="window.open($('#logo_faixa').val());" width="20" height="20" align="absmiddle"></td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_tec">&nbsp;</td>
                <td class="titulo_noticias_home"><div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:left;"></div>
                  <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco cursor" onClick="Enviar();">Atualizar Texto</div>
                  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:right; width:15px; height:17px; float:left;"></div></td>
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
      </table>
      <table width="900" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table></td>
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