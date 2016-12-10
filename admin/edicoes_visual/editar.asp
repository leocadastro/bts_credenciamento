<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
id	= Limpar_Texto(Request("id"))

'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")

	SQL_Eventos = 	"Select " &_
					"	Ee.ID_Evento, " &_
					"	Ee.ID_Edicao, " &_
					"	E.Nome_PTB as Evento, " &_
					"	Ee.Ano " &_
					"From Eventos_Edicoes as Ee " &_
					"Inner Join Eventos as E ON E.ID_Evento = Ee.ID_Evento " &_
					"Order by Ano DESC, Evento"
					
	Set RS_Eventos = Server.CreateObject("ADODB.Recordset")
	RS_Eventos.Open SQL_Eventos, Conexao

	SQL_Listar = 	"Select " &_
					"	Ecv.*, " &_
					"	E.Nome_PTB as Evento, " &_
					"	Ee.Ano " &_
					"From Edicoes_Configuracao as Ecv " &_
					"Inner Join Eventos_Edicoes as Ee ON Ee.ID_Edicao = Ecv.ID_Edicao " &_
					"Inner Join Eventos as E ON E.ID_Evento = Ee.ID_Evento " &_
					"Where ID_Configuracao = " & id

'	response.write("<hr>" & SQL_Listar & "<hr>")

	Set RS_Listar = Server.CreateObject("ADODB.Recordset")
	RS_Listar.Open SQL_Listar, Conexao, 1, 1
	
	If RS_Listar.BOF or RS_Listar.EOF Then
		response.Redirect("default.asp?msg=erro_nao_encontrado")
	Else
		evento 			= RS_Listar("evento")
		ano 			= RS_Listar("ano")
		ID_Configuracao	= RS_Listar("ID_Configuracao")
		id_edicao 		= RS_Listar("id_edicao")
		Cor		 		= RS_Listar("Cor")
		Faixa_Fundo 	= RS_Listar("Faixa_Fundo")
		Logo_Negativo 	= RS_Listar("Logo_Negativo")
		Logo_Box 		= RS_Listar("Logo_Box")
		Logo_Email 		= RS_Listar("Logo_Email")
		Template_Email 		= RS_Listar("Template_Email")
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
//Abre a pagina de Upload da Imagem
function UploadImagem(coluna)
{
    var id_edicao = document.cad.id_edicao;

    if (id_edicao.value == "")
    {
        alert("Selecine uma edição para fazer o upload da imagem!");
        id_edicao.focus();
        return;
    }

    var upimg = window.open("upload.asp?edicao=" + id_edicao.value + "&coluna=" + coluna, "Upload Imagem", "width=450, height=250");
    upimg.focus();
}
//Nova função para visualizar as imagens anexadas
function VisualizarImagem(coluna)
{
    var campo = document.cad[coluna];
    
    if (campo.value.indexOf(".") != -1)
    {
        var winimg = window.open(campo.value, "", "width=500, height=200, resizable=1");
        winimg.focus();
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
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">Configura&ccedil;&atilde;o Visual - <%=id_confVisual%><br><%=ano%> - <%=evento%></span></td>
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
            <input type="hidden" id="acao" name="acao" value="upd_visual">
            <input type="hidden" id="id" name="id" value="<%=id%>">
              <tr>
                <td height="30" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">Atualizar Configura&ccedil;&atilde;o Visual da Edi&ccedil;&atilde;o - <%=id_edicao%></span></td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_bts">ID Edi&ccedil;&atilde;o</td>
                <td class="titulo_noticias_home">
                <select id="id_edicao" name="id_edicao" class="admin_txtfield_login">
                <%
				If not RS_Eventos.BOF or not RS_Eventos.EOF Then
					While not RS_Eventos.EOF
					selecionado = ""
					If Cstr(id_edicao) = Cstr(RS_Eventos("ID_Edicao")) Then selecionado = " selected "
					%><option value="<%=RS_Eventos("ID_Edicao")%>" <%=selecionado%> ><%=RS_Eventos("ano")%> - <%=RS_Eventos("Evento")%></option><%
						RS_Eventos.MoveNext
					Wend
					RS_Eventos.Close
				End If
				%>
                </select>
                </td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Cor do Fundo</td>
                <td class="titulo_noticias_home">
                  <input name="cor_fundo" id="cor_fundo" type="text" class="admin_txtfield_login" size="10" style="float:left;" value="<%=cor%>" />
                  <div id="bg_cor_fundo" style="width:64px; height:16px; background-color:#<%=Replace(cor,"#","")%>; border:1px #000 solid; float:left; margin-left:10px;">&nbsp;</div>
                </td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Faixa do Fundo</td>
                <td class="titulo_noticias_home">
                  <input name="faixa_fundo" id="faixa_fundo" type="text" class="admin_txtfield_login" size="30"  value="<%=faixa_fundo%>" />
                    <img src="../images/ico_upload.png" alt="Upload Imagem" title="Upload Imagem" border="0" width="20" align="middle" onclick="UploadImagem('faixa_fundo')" class="cursor" />
                    <img src="../images/ico_preview_20_b.gif" alt="Visualizar Imagem" title="Visualizar Imagem" border="0" width="20" align="middle" onclick="VisualizarImagem('faixa_fundo')" class="cursor" />
                </td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Logotipo da Faixa</td>
                <td class="titulo_noticias_home">
                  <input name="logo_faixa" id="logo_faixa" type="text" class="admin_txtfield_login" size="30"  value="<%=logo_negativo%>" />
                    <img src="../images/ico_upload.png" alt="Upload Imagem" title="Upload Imagem" border="0" width="20" align="middle" onclick="UploadImagem('logo_faixa')" class="cursor" />
                    <img src="../images/ico_preview_20_b.gif" alt="Visualizar Imagem" title="Visualizar Imagem" border="0" width="20" align="middle" onclick="VisualizarImagem('logo_faixa')" class="cursor" />
                </td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Logotipo do Box</td>
                <td class="titulo_noticias_home">
                  <input name="logo_box" id="logo_box" type="text" class="admin_txtfield_login" size="30"  value="<%=logo_box%>" />
                    <img src="../images/ico_upload.png" alt="Upload Imagem" title="Upload Imagem" border="0" width="20" align="middle" onclick="UploadImagem('logo_box')" class="cursor" />
                    <img src="../images/ico_preview_20_b.gif" alt="Visualizar Imagem" title="Visualizar Imagem" border="0" width="20" align="middle" onclick="VisualizarImagem('logo_box')" class="cursor" />
                </td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Logotipo para Email de Confirm.</td>
                <td class="titulo_noticias_home"><input name="logo_email" id="logo_email" type="text" class="admin_txtfield_login" size="30"  value="<%=logo_email%>" />
                    <img src="../images/ico_upload.png" alt="Upload Imagem" title="Upload Imagem" border="0" width="20" align="middle" onclick="UploadImagem('logo_email')" class="cursor" />
                    <img src="../images/ico_preview_20_b.gif" alt="Visualizar Imagem" title="Visualizar Imagem" border="0" width="20" align="middle" onclick="VisualizarImagem('logo_email')" class="cursor" />
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_tec"><span class="titulo_menu_site_bts">Template do Email de Confirm.</span></td>
                <td class="titulo_noticias_home"><input name="url_template" id="url_template" type="text" class="admin_txtfield_login" size="30"  value="<%=template_email%>" />
                    <img src="../images/ico_upload.png" alt="Upload Imagem" title="Upload Imagem" border="0" width="20" align="middle" onclick="UploadImagem('url_template')" class="cursor" />
                    <img src="../images/ico_preview_20_b.gif" alt="Visualizar Imagem" title="Visualizar Imagem" border="0" width="20" align="middle" onclick="VisualizarImagem('url_template')" class="cursor" />
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