<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
id	= Limpar_Texto(Request("id"))

'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")

	SQL_Idioma = 	"Select " &_
					"	ID_Idioma, " &_
					"	Nome " &_
					"From Idiomas " &_
					"Where Ativo = 1 " &_
					"Order by ID_Idioma"
					
	Set RS_Idioma = Server.CreateObject("ADODB.Recordset")
	RS_Idioma.Open SQL_Idioma, Conexao

	SQL_Listar = 	"Select " &_
					"	ID_Idioma,  " &_
					"	Nome,  " &_
					"	IMG_Faixa,  " &_
					"	IMG_Box,  " &_
					"	URL,  " &_
					"	Ativo,  " &_
					"	Data_Cadastro as Data " &_
					"From Tipo_Credenciamento " &_
					"Where ID_Tipo_Credenciamento = " & id

'	response.write("<hr>" & SQL_Listar & "<hr>")

	Set RS_Listar = Server.CreateObject("ADODB.Recordset")
	RS_Listar.Open SQL_Listar, Conexao
	
	If RS_Listar.BOF or RS_Listar.EOF Then
		response.Redirect("default.asp?msg=erro_nao_encontrado")
	Else
		ID_Idioma	= RS_Listar("ID_Idioma")
		Nome		= RS_Listar("Nome")
		IMG_Faixa	= RS_Listar("IMG_Faixa")
		IMG_Box	= RS_Listar("IMG_Box")
		URL			= RS_Listar("URL")
		Ativo		= RS_Listar("Ativo")
		RS_Listar.Close
	End If
%>
<html>
<head>
<meta http-equiv="Content-type" content="text/html; charset=iso-8859-1" />
<title>Administração CSC - Brazil Trade Shows</title>
<link href="/admin/css/bts.css" rel="stylesheet" type="text/css" />
<link href="/admin/css/admin.css" rel="stylesheet" type="text/css" />
<link href="/css/colorpicker.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
<script language="javascript" src="/admin/js/validar_forms.js"></script>
<script language="javascript" src="/js/colorpicker/colorpicker.js"></script>
</head>

<script language="javascript">
$(document).ready(function(){
	$('#aviso').hide();
	
	$('#cor_fundo').ColorPicker({
		onSubmit: function(hsb, hex, rgb, el) {
			$(el).val('#' + hex);
			$('#bg_cor_fundo').css('background-color','#' + hex);
			$(el).ColorPickerHide();
		},
		onBeforeShow: function() {
			$(this).ColorPickerSetColor(this.value);
		}
	});
	
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
			case 'id':
				break;
			case 'acao':
				break;
			default:
				if (this.id.lenght > 0) {
					erros += verificar(this.id, false)
				}
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
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">Tipo de Crenciamento</span></td>
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
            <form action="metodos.asp" method="post" name="cad" id="cad">
            <input type="hidden" id="acao" name="acao" value="upd_tipo_cred">
            <input type="hidden" id="id" name="id" value="<%=id%>">
              <tr>
                <td height="30" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">Atualizar</span></td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_bts">Idioma</td>
                <td class="titulo_noticias_home">
                <select id="id_idioma" name="id_idioma" class="admin_txtfield_login">
                <option value="-">-- Selecione --</option>
                <%
				If not RS_Idioma.BOF or not RS_Idioma.EOF Then
					While not RS_Idioma.EOF
						selecionado = ""
						If Cstr(ID_Idioma) = Cstr(RS_Idioma("ID_Idioma")) Then selecionado = " selected "
						%><option value="<%=RS_Idioma("ID_Idioma")%>" <%=selecionado%>><%=RS_Idioma("Nome")%></option><%
						RS_Idioma.MoveNext
					Wend
					RS_Idioma.Close
				End If
				%>
                </select>
                </td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_bts">Tipo do Credenciamento</td>
                <td class="titulo_noticias_home"><input name="nome" id="nome" type="text" class="admin_txtfield_login" size="30" value="<%=nome%>" /></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Imagem Faixa</td>
                <td class="titulo_noticias_home"><input name="img_faixa" id="img_faixa" type="text" class="admin_txtfield_login" size="30"  value="<%=img_faixa%>" />
                  <img src="/admin/images/ico_preview_20_b.gif" alt="Visualizar Imagem" title="Visualizar Imagem" class="cursor" onClick="window.open($('#img_faixa').val());" width="20" height="20" align="absmiddle">
                </td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Imagem Box</td>
                <td class="titulo_noticias_home"><input name="img_box" id="img_box" type="text" class="admin_txtfield_login" size="30"  value="<%=img_box%>" />
                  <img src="/admin/images/ico_preview_20_b.gif" alt="Visualizar Imagem" title="Visualizar Imagem" class="cursor" onClick="window.open($('#img_box').val());" width="20" height="20" align="absmiddle">
                </td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">URL do Formul&aacute;rio</td>
                <td class="titulo_noticias_home"><input name="url" id="url" type="text" class="admin_txtfield_login" size="30"  value="<%=url%>" /></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Dispon&iacute;vel</td>
                <td class="titulo_noticias_home">
                <select id="ativo" name="ativo" class="admin_txtfield_login">
                    <option value="1" <% If ativo = "1" OR ativo = true Then %> selected <% End If %> >Sim</option>
                    <option value="0" <% If ativo = "0" OR ativo = false Then %> selected <% End If %> >Não</option>
                </select>
                </td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_tec">&nbsp;</td>
                <td class="titulo_noticias_home">
                  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:left;"></div>
                  <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco cursor" onClick="Enviar();">Atualizar Configura&ccedil;&atilde;o</div>
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