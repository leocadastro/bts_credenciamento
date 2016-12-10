<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
id	= Limpar_Texto(Request("id"))

'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")

	SQL_Eventos = 	"Select " &_
					"	Ee.ID_Edicao, " &_
					"	E.Nome_PTB as Evento, " &_
					"	Ee.Ano " &_
					"From Eventos_Edicoes as Ee " &_
					"Inner Join Eventos as E ON E.ID_Evento = Ee.ID_Evento " &_
					"Order by Ano DESC, Evento"
					
	Set RS_Eventos = Server.CreateObject("ADODB.Recordset")
	RS_Eventos.Open SQL_Eventos, Conexao

	SQL_Tipos = 	"Select " &_
					"	Tc.ID_Tipo_Credenciamento,  " &_
					"	I.Nome as Idioma, " &_
					"	Tc.ID_Idioma, " &_
					"	Tc.Nome,  " &_
					"	Tc.URL  " &_
					"From Tipo_Credenciamento as Tc " &_
					"Inner Join Idiomas as I on I.ID_Idioma = Tc.ID_Idioma " &_
					"Order by Tc.ID_Idioma, Tc.Nome"
					
	Set RS_Tipos = Server.CreateObject("ADODB.Recordset")
	RS_Tipos.Open SQL_Tipos, Conexao

	SQL_Listar = 	"Select " &_
					"	ID_Edicao,  " &_
					"	ID_Tipo_Credenciamento,  " &_
					"	Inicio,  " &_
					"	Fim,  " &_
					"	URL_Especial,  " &_
					"	Ativo  " &_ 
					"From Edicoes_Tipo " &_
					"Where ID_Edicao_Tipo = " & ID

'	response.write("<hr>" & SQL_Listar & "<hr>")

	Set RS_Listar = Server.CreateObject("ADODB.Recordset")
	RS_Listar.Open SQL_Listar, Conexao
	
	If RS_Listar.BOF or RS_Listar.EOF Then
		response.Redirect("default.asp?msg=erro_nao_encontrado")
	Else
		ID_Edicao				= RS_Listar("ID_Edicao")
		ID_Tipo_Credenciamento	= RS_Listar("ID_Tipo_Credenciamento")
		Inicio					= RS_Listar("Inicio")
		Fim						= RS_Listar("Fim")
			
		data_ini = Replace(Left(Inicio,10),"/",".")
        hora_ini = Mid(Inicio,12,16)
		If Len(Trim(hora_ini)) = 0 Then hora_ini = "00:00"
		
		data_fim = Replace(Left(Fim,10),"/",".")
        hora_fim = Mid(Fim,12,16)
		
		URL_Especial			= RS_Listar("URL_Especial")
		Ativo					= RS_Listar("Ativo")
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
<link href="/css/calendar.css" rel="stylesheet" type="text/css" media="screen">

<script language="javascript" src="/admin/ckeditor/ckeditor.js"></script>
<script language="javascript" src="/admin/ckeditor/contents.css"></script> 

<script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
<script language="javascript" src="/js/jquery.maskedinput-1.2.2.min.js"></script>
<script language="javascript" src="/admin/js/validar_forms.js"></script>
<script language="javascript" src="/js/colorpicker/colorpicker.js"></script>
<script language="javascript" src="/js/Calendario/calendar.js"></script>
</head>

<script language="javascript">
$(document).ready(function(){
	
	
	$('#aviso').hide();
	$('#hora_ini').mask("99:99",{placeholder:"_"});
	$('#hora_fim').mask("99:99",{placeholder:"_"});
	
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
	
	//atribui valores da mono para o URL_Especial
	$('#url_especial').value = $('iframe').contents().find('body').html();
	$('#url_especial').attr('value', $('iframe').contents().find('body').html())


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
			case 'url_especial':
				break;
			default:
				erros += verificar(this.id, false)
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
            <input type="hidden" id="acao" name="acao" value="upd_rel_evt_tipo">
            <input type="hidden" id="id" name="id" value="<%=id%>">
              <tr>
                <td height="30" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">Atualizar</span></td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_bts">Edição</td>
                <td class="titulo_noticias_home">
                <select id="id_edicao" name="id_edicao" class="admin_txtfield_login">
                <option value="-">-- Selecione --</option>
                <%
				If not RS_Eventos.BOF or not RS_Eventos.EOF Then
					While not RS_Eventos.EOF
						selecionado = ""
						If Cstr(ID_Edicao) = Cstr(RS_Eventos("ID_Edicao")) Then selecionado = " selected "
						%><option value="<%=RS_Eventos("ID_Edicao")%>" <%=selecionado%>><%=RS_Eventos("Ano")%> - <%=RS_Eventos("Evento")%></option><%
						RS_Eventos.MoveNext
					Wend
					RS_Eventos.Close
				End If
				%>
                </select>
                </td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_bts">Tipo do Credenciamento</td>
                <td class="titulo_noticias_home">
                <select id="id_tipo" name="id_tipo" class="admin_txtfield_login">
                <option value="-">-- Selecione --</option>
                <%
				If not RS_Tipos.BOF or not RS_Tipos.EOF Then
					idioma_listado = ""
					While not RS_Tipos.EOF
						If Cstr(idioma_listado) <> Cstr(RS_Tipos("ID_Idioma")) Then
							idioma_listado = RS_Tipos("ID_Idioma")
							%><optgroup label="• <%=RS_Tipos("Idioma")%>"><%
						End If
						
						selecionado = ""
						If Cstr(ID_Tipo_Credenciamento) = Cstr(RS_Tipos("ID_Tipo_Credenciamento")) Then selecionado = " selected "
						%><option value="<%=RS_Tipos("ID_Tipo_Credenciamento")%>" <%=selecionado%>><%=RS_Tipos("Nome")%> - <%=RS_Tipos("URL")%></option><%
						
						RS_Tipos.MoveNext
						If not RS_Tipos.EOF Then
							If Cstr(idioma_listado) <> Cstr(RS_Tipos("ID_Idioma")) Then
								%></optgroup><%
							End If
						End IF
					Wend
					RS_Tipos.Close
				End If
				%>
                </select>
                </td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Per&iacute;odo para Cadastros : In&iacute;cio</td>
                <td class="t_arial fs11px bold c_vermelho">
                  <input name="data_ini" id="data_ini" type="text" size="12" class="admin_txtfield_login" readonly value="<%=data_ini%>">
                  <img src="/admin/images/img_calendario.gif" width="28" height="24" border="0" align="absmiddle" onClick="displayCalendar(document.forms[0].data_ini,'dd.mm.yyyy',this);" class="bt_aba" id="img_calendario1" />
                  às <input name="hora_ini" id="hora_ini" type="text" size="6" class="admin_txtfield_login" value="<%=hora_ini%>">
                </td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Per&iacute;odo para Cadastros : Fim</td>
                <td class="t_arial fs11px bold c_vermelho">
                  <input name="data_fim" id="data_fim" type="text" size="12" class="admin_txtfield_login" readonly value="<%=data_fim%>">
                  <img src="/admin/images/img_calendario.gif" width="28" height="24" border="0" align="absmiddle" onClick="displayCalendar(document.forms[0].data_fim,'dd.mm.yyyy',this);" class="bt_aba" id="img_calendario1" />
                  às <input name="hora_fim" id="hora_fim" type="text" size="6" class="admin_txtfield_login" value="<%=hora_fim%>">
                </td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts" id="titulo_menu_site_bts">Formul&aacute;rio Especial ? URL :</td>
                <td class="titulo_noticias_home">
                </td>
              </tr>
              <tr>
              	<td colspan="2">
                <textarea name="url_especial" id="url_especial" class="admin_txtfield_login" /><%=URL_Especial%></textarea>
				<script type="text/javascript">
                //<![CDATA[
                CKEDITOR.replace('url_especial',{
                    skin : 'moono',
					appendTo:'',
					removePlugins : 'flash,about',
                    htmlEncodeOutput : true,
                    entities : false
                });
									
				</script>
                </td>
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