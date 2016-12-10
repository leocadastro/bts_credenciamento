<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
' * Dados Paginação
qtde					= Limpar_Texto(Request("qtd"))
ordem					= Limpar_Texto(Request("ord"))
acao_pagina				= Limpar_Texto(Request("acp"))
pagina					= Limpar_Texto(Request("pag"))

strURL = "default.asp?w=0"

If Len(qtde) > 0 AND isNumeric(qtde) Then 
	qtde_itens = qtde
	strURL = strURL & "&qtd=" & qtde
Else
	qtde_itens = 200
	qtde = 200
End If

'==================================================
dia = Day(Now)
If Len(dia) < 2 Then dia = "0" & dia
mes = Month(Now)
If Len(mes) < 2 Then mes = "0" & mes
ano = Year(Now)
hoje = dia & "." & mes & "." & ano
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
	
	SQL_Eventos = 	"Select " &_
					"	Ee.ID_Edicao, " &_
					"	E.Nome_PTB as Evento, " &_
					"	Ee.Ano " &_
					"From Eventos_Edicoes as Ee " &_
					"Inner Join Eventos as E ON E.ID_Evento = Ee.ID_Evento " &_
					"Inner Join Edicoes_Configuracao as Ec ON Ec.ID_Edicao = Ee.ID_Edicao " &_
					"Where " &_
					"	Ee.Ativo = 1 " &_
					"	AND Ec.Ativo = 1 " &_
					"	AND E.Ativo = 1 " &_
					"Order by Ano DESC, Evento"
	'response.write("SQL_Eventos: " & SQL_Eventos & "<br/>")
					
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
					"	Et.ID_Edicao_Tipo,  " &_
					"	Et.ID_Edicao, " &_
					"	E.Nome_PTB as Evento,  " &_
					"	Ee.Ano,  " &_
					"	Ec.Cor, " &_
					"	Ec.Logo_Box, " &_
					"	Tc.ID_Idioma, " &_
					"	Tc.Nome as Tipo_Credenciamento,  " &_
					"	Tc.IMG_Box, " &_
					"	Et.Inicio,  " &_
					"	Et.Fim,  " &_
					"	Et.URL_Especial,  " &_
					"	Et.Ativo,  " &_
					"	Et.Data_Cadastro as Data  " &_
					"From Edicoes_Tipo as Et  " &_
					"Inner Join Eventos_Edicoes as Ee ON Ee.ID_Edicao = Et.ID_Edicao  " &_
					"Inner Join Edicoes_Configuracao as Ec ON Ec.ID_Edicao = Et.ID_Edicao " &_
					"Inner Join Eventos as E ON E.ID_Evento = Ee.ID_Evento  " &_
					"Inner Join Tipo_Credenciamento as Tc ON Tc.ID_Tipo_Credenciamento = Et.ID_Tipo_Credenciamento  " &_
					"Where " &_
					"	Ee.Ativo = 1 " &_
					"	AND Ec.Ativo = 1 " &_
					"	AND E.Ativo = 1 " &_
					"	AND Tc.Ativo = 1 " &_
					"	AND Ee.Ano >= Year(getDate()) " &_
					"Order by Et.ID_Edicao DESC, Tc.ID_Idioma, Tc.Nome"

'	response.write("<hr>" & SQL_Listar & "<hr>")

	Set RS_Listar = Server.CreateObject("ADODB.Recordset")
		RS_Listar.CursorLocation = 3
		RS_Listar.CacheSize = qtde_itens
		RS_Listar.PageSize = qtde_itens
	RS_Listar.Open SQL_Listar, Conexao, 3, 3
	
	If not RS_Listar.BOF and not RS_Listar.EOF Then
		TotalPaginas = RS_Listar.PageCount	
	Else 
		TotalPaginas = 0 
	End IF
'==================================================
intPageCount = TotalPaginas
Select Case acao_pagina
		Case "I" ' inicio
			intpage = 1
		Case "a" ' anterior
			intpage = pagina - 1
			if intpage < 1 then intpage = 1
		Case "p" ' proxima
			intpage = pagina + 1
			IF intpage > intPageCount Then intpage = IntPageCount
		Case "U" 'ultima
			intpage = intPageCount
		Case "n" ' numero X
			intpage = pagina
		Case Else
			intpage = 1
End Select
'==================================================
PaginaAtual = intpage
'==================================================
If paginacao = "" then
	If Request.QueryString("pg") <> "" Then
		PaginaAtual = Cint(Request.QueryString("pg"))
		intpage = PaginaAtual
	End If
End If
'==================================================
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
			$('#aviso_conteudo').html('Adicionado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
			<%
		Case "add_erro_existe"
			%>
			$('#aviso_conteudo').html('Erro - Item informado já existe !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
			<%
		Case "upd_ok"
			%>
			$('#aviso_conteudo').html('Atualizado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
			<%
		Case "atv_ok"
			%>
			$('#aviso_conteudo').html('Ativado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
			<%
		Case "des_ok"
			%>
			$('#aviso_conteudo').html('Desativado !');
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
			default:
				if (this.id.lenght > 0) {
					erros += verificar(this.id, false);
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
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center">
        <span style="color: #B01D22">
        	Relacionar:<br>
	        Eventos > Tipos de Credenciamento
        </span>
		</td>
        <td width="100" align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;"><a href="/admin/menu.asp"><img src="/admin/images/bt_voltar_admin.gif" width="100" height="48" border="0" /></a></td>
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
            <input type="hidden" id="acao" name="acao" value="add_rel_evt_tipo">
              <tr>
                <td height="30" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">Cadastrar</span></td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_bts">Edição</td>
                <td class="titulo_noticias_home">
                <select id="id_edicao" name="id_edicao" class="admin_txtfield_login">
                <option value="-">-- Selecione --</option>
                <%
				If not RS_Eventos.BOF or not RS_Eventos.EOF Then
					While not RS_Eventos.EOF
					%><option value="<%=RS_Eventos("ID_Edicao")%>"><%=RS_Eventos("Ano")%> - <%=RS_Eventos("Evento")%></option><%
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
						
						%><option value="<%=RS_Tipos("ID_Tipo_Credenciamento")%>"><%=RS_Tipos("Nome")%> - <%=RS_Tipos("URL")%></option><%
						
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
                  <input name="data_ini" id="data_ini" type="text" size="12" class="admin_txtfield_login" readonly value="<%=hoje%>">
                  <img src="/admin/images/img_calendario.gif" width="28" height="24" border="0" align="absmiddle" onClick="displayCalendar(document.forms[0].data_ini,'dd.mm.yyyy',this);" class="bt_aba" id="img_calendario1" />
                  às <input name="hora_ini" id="hora_ini" type="text" size="6" class="admin_txtfield_login" value="00:00">
                </td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Per&iacute;odo para Cadastros : Fim</td>
                <td class="t_arial fs11px bold c_vermelho">
                  <input name="data_fim" id="data_fim" type="text" size="12" class="admin_txtfield_login" readonly value="<%=hoje%>">
                  <img src="/admin/images/img_calendario.gif" width="28" height="24" border="0" align="absmiddle" onClick="displayCalendar(document.forms[0].data_fim,'dd.mm.yyyy',this);" class="bt_aba" id="img_calendario1" />
                  às <input name="hora_fim" id="hora_fim" type="text" size="6" class="admin_txtfield_login" value="23:59">
                </td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Formul&aacute;rio Especial ? URL :</td>
                <td class="titulo_noticias_home">
                </td>
              </tr>
              <tr>
              	<td colspan="2">
				<input name="url" id="url" type="text" class="admin_txtfield_login" size="30"  />
				<script type="text/javascript">
                //<![CDATA[
                CKEDITOR.replace('url',{
                    skin : 'moono',
                    removePlugins : 'save,flash,about',
                    htmlEncodeOutput : false,
                    entities : false
                });
                
                //]]>
                </script>               
                <br>
                </td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Dispon&iacute;vel</td>
                <td class="titulo_noticias_home">
                <select id="ativo" name="ativo" class="admin_txtfield_login">
                  <option value="0">N&atilde;o</option>
                  <option value="1">Sim</option>
                </select>
                </td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_tec">&nbsp;</td>
                <td class="titulo_noticias_home">
                  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:left;"></div>
                  <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco cursor" onClick="Enviar();">Cadastrar Nova Configura&ccedil;&atilde;o</div>
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
      <table width="900" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table>
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
			<%	If Rs_listar.BOF or Rs_listar.EOF Then	%>
            <p align="center" class="titulo_menu_site_carne">N&atilde;o foi encontrado nenhum registro</p>
            <% End If %>
            <%
            Contador = 0
            If not Rs_listar.BOF or not Rs_listar.EOF Then
            %>
                <%
                Rs_listar.MoveFirst
                RS_Listar.AbsolutePage = PaginaAtual 
				
				feira_exibida = ""
				idioma_exibido = ""
                While Not RS_Listar.EOF And Contador < RS_Listar.PageSize
                    Contador = Contador + 1
					
					ID_Edicao_Tipo		= RS_Listar("ID_Edicao_Tipo")
					ID_Edicao			= RS_Listar("ID_Edicao")
					Evento				= RS_Listar("Evento")
					Ano					= RS_Listar("Ano")
					Cor					= RS_Listar("Cor")
					Logo_Box			= RS_Listar("Logo_Box")
					ID_Idioma			= RS_Listar("ID_Idioma")
					Tipo_Credenciamento	= RS_Listar("Tipo_Credenciamento")
					IMG_Box				= RS_Listar("IMG_Box")
					Inicio				= RS_Listar("Inicio")
					Fim					= RS_Listar("Fim")
					URL_Especial		= RS_Listar("URL_Especial")
					Ativo				= RS_Listar("Ativo")
					Data				= RS_Listar("Data")
                    
					Select Case ID_Idioma
						Case "1"
							sigla_idioma = "ptb"
						Case "2"
							sigla_idioma = "esp"
						Case "3"
							sigla_idioma = "eng"
					End Select
					
                    If Cstr(feira_exibida) <> Cstr(ID_Edicao) Then
						feira_exibida = ID_Edicao
						idioma_exibido = ""
						%>
						<!-- Cabecalho FEIRA -->
                        <br>
						<table width="900" border="0" cellspacing="0" cellpadding="0" align="center">
						  <tr>
							<td height="4" bgcolor="<%=cor%>"><img src="/img/geral/spacer.gif" width="110" height="4"></td>
						  </tr>
						  <tr>
							<td height="50" align="center"><img src="<%=Logo_Box%>" border="0"></td>
						  </tr>
						</table>
	                    <!-- Cabecalho FEIRA -->
                        <%
					End If
					
					If Cstr(idioma_exibido) <> CStr(ID_Idioma) Then
						idioma_exibido = ID_Idioma
					%>
                    <!-- Bandeira e Tipos -->
                    <hr size="1">
                    <table width="900" border="0" cellspacing="0" cellpadding="0" align="center">
                      <tr>
                        <td width="50"><img src="/img/geral/bandeira_<%=sigla_idioma%>.gif" width="48" height="29"></td>
                        <td>
                    <% End If %>
                        <div style="width:195px; background-color:#fff; float:left; padding-left:10px; padding-top:10px;">
                            <div style="width:195px; height:4px; background-color:#5a5a5a;"><img src="/img/geral/spacer.gif" width="110" height="4"></div>
                          <div style="width:195px; padding:5px;" class="fs12px t_arial">
                                <img src="<%=img_box%>" border="0" hspace="20" vspace="5"><br>
                                <b>Início:</b>&nbsp;<%=inicio%><br>
                                <b>Fim:</b>&nbsp;<%=fim%><br>
                                <b>Formulário Esp.:</b>&nbsp;<%=url_especial%><br>
                                <b>Disponível:</b>&nbsp;
								<%
                                If ativo = true OR ativo = 1 Then
                                    %><img src="/img/geral/icones/ok.gif" width="20" height="20" alt="ativo" title="ativo" onClick="document.location = 'metodos.asp?id=<%=Rs_listar("ID_Edicao_Tipo")%>&acao=desativar';" class="cursor"><%
                                Else
                                    %><img src="/img/geral/icones/nok.gif" width="20" height="20" alt="desativado" title="desativado" onClick="document.location = 'metodos.asp?id=<%=Rs_listar("ID_Edicao_Tipo")%>&acao=ativar';" class="cursor"><%
                                End If
                                %><br>
                            Editar: <a href="editar.asp?id=<%=Rs_listar("ID_Edicao_Tipo")%>"><img src="/admin/images/ico_editar.gif" width="20" height="20"></a>
                            </div>
                            <div style="width:195px; height:4px; background-color:#5a5a5a;"><img src="/img/geral/spacer.gif" width="110" height="4"></div>
                        </div>
            		<%

                    RS_Listar.MoveNext()
					If not RS_Listar.EOF Then
						If Cstr(idioma_exibido) <> CStr(RS_Listar("ID_Idioma")) Then
						%>
							</td>
						  </tr>
						</table>
						<%
						End If
					Else
						%>
							</td>
						  </tr>
						</table>
						<%
					End If
                Wend
                RS_Listar.Close
            End If
            %>
          <!--#include virtual="/admin/inc/paginacao.asp"-->
          </td>
          <td background="/admin/images/caixa/dir.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12"></td>
        </tr>
        <tr>
          <td><img src="/admin/images/caixa/inf_esq.gif" width="14" height="12"></td>
          <td background="/admin/images/caixa/inf_centro.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12"></td>
          <td><img src="/admin/images/caixa/inf_dir.gif" width="14" height="12"></td>
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