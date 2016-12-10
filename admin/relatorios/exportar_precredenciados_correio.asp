<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
' * Dados Paginação
qtde					= Limpar_Texto(Request("qtd"))
ordem					= Limpar_Texto(Request("ord"))
acao_pagina				= Limpar_Texto(Request("acp"))
pagina					= Limpar_Texto(Request("pag"))
id_edicao 				= Limpar_Texto(Request("id_edicao"))

strURL = "?w=0"

If Len(qtde) > 0 AND isNumeric(qtde) Then 
	qtde_itens = qtde
	strURL = strURL & "&qtd=" & qtde
Else
	qtde_itens = 40
	qtde = 40
End If

'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")

  SQL_Eventos =   "Select " &_
          " Ee.ID_Evento, " &_
          " Ee.ID_Edicao, " &_
          " E.Nome_PTB as Evento, " &_
          " Ee.Ano " &_
          "From Eventos_Edicoes as Ee " &_
          "Inner Join Eventos as E ON E.ID_Evento = Ee.ID_Evento " &_
          "Order by Ano DESC, Evento"
  Set RS_Eventos = Server.CreateObject("ADODB.Recordset")
  RS_Eventos.Open SQL_Eventos, Conexao
  
  SQL_Idiomas =   "Select " &_
          " ID_Idioma " &_
		  " ,Nome " &_
		  " ,Ativo " &_
          "From Idiomas " &_
          "Order by ID_Idioma"
  Set RS_Idiomas = Server.CreateObject("ADODB.Recordset")
  RS_Idiomas.Open SQL_Idiomas, Conexao
  
  SQL_Tipos =   "Select " &_
          " ID_Formulario " &_
		  " ,Nome " &_
          "From Formularios " &_
          "Order by ID_Formulario"
  Set RS_Tipos = Server.CreateObject("ADODB.Recordset")
  RS_Tipos.Open SQL_Tipos, Conexao
  
  If Len(Trim(ID_Edicao)) > 0 Then
	SQL_Arquivos = 	"Select " &_
					"	Arquivo " &_ 
					"	,Total " &_ 
					"	,Data_Cadastro " &_
					"From Arquivos_XML " &_ 
					"Where  " &_ 
					"	Ativo = 1 " &_ 
					"	AND ID_Edicao = " & id_edicao & " " &_
					"Order by Data_Cadastro DESC "
	Set RS_Arquivos = Server.CreateObject("ADODB.Recordset")
	RS_Arquivos.Open SQL_Arquivos, Conexao
  End If
  
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
<title>Administração Cred. 2012</title>
<link href="/admin/css/bts.css" rel="stylesheet" type="text/css" />
<link href="/admin/css/admin.css" rel="stylesheet" type="text/css" />
<link href="/css/calendar.css" rel="stylesheet" type="text/css" media="screen">
<script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
<script language="javascript" src="/admin/js/validar_forms.js"></script>
<script language="javascript" src="/js/colorpicker/colorpicker.js"></script>
<script language="javascript" src="/js/Calendario/calendar.js"></script>
<script language="javascript" src="/js/jquery.maskedinput-1.2.2.min.js"></script>
</head>


<script language="javascript">
$(document).ready(function(){
	$('#aviso').hide();
	$('#hora_ini').mask("99:99",{placeholder:"_"});

	<% 
	msg = Request("msg")
	Select Case msg
		Case "pag_proibida"
			Session("admin_msg") = ""
			%>
			$('#aviso_conteudo').html('Página não permitida !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 10000);
			<%
		Case "add_ok"
			%>
			$('#aviso_conteudo').html('Item adicionado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 10000).fadeOut('slow');
			<%
		Case "erro_nao_encontrado"
			%>
			$('#aviso_conteudo').html('Erro - Não foi encontrado nenhum registro !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 10000).fadeOut('slow');
			<%
		Case "add_erro_existe"
			%>
			$('#aviso_conteudo').html('Erro - Item informado já existe !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 10000).fadeOut('slow');
			<%
		Case "upd_ok"
			%>
			$('#aviso_conteudo').html('Item atualizado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 10000).fadeOut('slow');
			<%
		Case "atv_ok"
			%>
			$('#aviso_conteudo').html('Item ativado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 10000).fadeOut('slow');
			<%
		Case "des_ok"
			%>
			$('#aviso_conteudo').html('Item desativado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 10000).fadeOut('slow');
			<%
	End Select
	%>
});

function Enviar() {
	var erros = 0;
	$('select:enabled').each(function(i) {
		// Se não for obrigatório
		switch (this.id) {
			case "id_idioma":
				break;
			case "id_tipo":
				break;
			default:
				erros += verificar(this.id, false);
				break;
		}
	});
	if (erros == 0) {
		document.buscar.submit();	
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
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">Exportar Pré-Credenciados - POSTAGEM CORREIO</span></td>
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
            <form id="buscar" name="buscar" method="get" target="_blank" action="exportar_correio.asp?busca=exportar">
              <tr>
                <td height="30" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">Exportar Pré-Credenciado</span></td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_bts"> Edi&ccedil;&atilde;o</td>
                <td class="titulo_noticias_home">
                <select id="id_edicao" name="id_edicao" class="admin_txtfield_login" onChange="document.location = '?id_edicao=' + this.value">
                <option value="">-- Selecione --</option>
                <%
                  If not RS_Eventos.BOF or not RS_Eventos.EOF Then
                    While not RS_Eventos.EOF
						selecionado = ""
						If Cstr(id_edicao) = Cstr(RS_Eventos("ID_Edicao")) Then selecionado = " selected "
    	                %><option value="<%=RS_Eventos("ID_Edicao")%>" <%=selecionado%>><%=RS_Eventos("ano")%> - <%=RS_Eventos("Evento")%></option><%
                      RS_Eventos.MoveNext
                    Wend
                    RS_Eventos.Close
                  End If
                  %>
                </select>
                </td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Tipo</td>
                <td class="titulo_noticias_home">
                <select id="id_tipo" name="id_tipo" class="admin_txtfield_login">
                	<option value="">-- Selecione --</option>
                <%
                  If not RS_Tipos.BOF or not RS_Tipos.EOF Then
                    While not RS_Tipos.EOF
                    %><option value="<%=RS_Tipos("ID_Formulario")%>"><%=RS_Tipos("Nome")%></option><%
                      RS_Tipos.MoveNext
                    Wend
                    RS_Tipos.Close
                  End If
                  %>
                </select>
                </td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Idioma</td>
                <td class="titulo_noticias_home">
                <select id="id_idioma" name="id_idioma" class="admin_txtfield_login">
                    <option value="">-- Selecione --</option>
                <%
                  If not RS_Idiomas.BOF or not RS_Idiomas.EOF Then
                    While not RS_Idiomas.EOF
                    %><option value="<%=RS_Idiomas("ID_Idioma")%>"><%=RS_Idiomas("Nome")%></option><%
                      RS_Idiomas.MoveNext
                    Wend
                    RS_Idiomas.Close
                  End If
                  %>
                </select>
                </td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_tec">&nbsp;</td>
                <td class="titulo_noticias_home">
                  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:left;"></div>
                  <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco cursor" onClick="Enviar()">Gerar Relatório em XLS</div>
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
      </td>
    </tr>
<%   If Len(Trim(ID_Edicao)) > 0 Then %>
  <tr>
    <td align="center" bgcolor="#FFFFFF">
	<br>
    <br>
    <table width="90%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999" class="conteudo_site">
    <%
      If RS_Arquivos.BOF or RS_Arquivos.EOF Then
        %>
        <tr>
            <td colspan="3" bgcolor="#FFFFFF" align="center">Não foram encontrados arquivos gerados anteriormente !</td>
        </tr>
        <%
      ElseIf not RS_Arquivos.BOF or not RS_Arquivos.EOF Then
        %>
          <tr>
            <td width="50" align="center" bgcolor="#FFFFFF" class="conteudo_home"><strong>N&ordm;</strong></td>
            <td align="center" bgcolor="#FFFFFF" class="conteudo_home"><strong>Registros</strong></td>
            <td align="center" bgcolor="#FFFFFF" class="conteudo_home"><strong>Arquivo</strong></td>
            <td align="center" bgcolor="#FFFFFF" class="conteudo_home"><strong>Data</strong></td>
          </tr>
        <%
        n = 0
        While not RS_Arquivos.EOF
            n = n + 1
            %>
            <tr>
                <td height="25" bgcolor="#FFFFFF" align="center"><%=n%></td>
                <td bgcolor="#FFFFFF" align="center"><%=RS_Arquivos("total")%></td>
                <td bgcolor="#FFFFFF" align="center"><a href="/admin/exportar_xml/arquivos_2012/<%=RS_Arquivos("arquivo")%>" target="_blank"><%=RS_Arquivos("arquivo")%></a></td>
                <td bgcolor="#FFFFFF" align="center"><%=FormatDateTime(RS_Arquivos("data_cadastro"),2)%></td>
            </tr>
            <%
            RS_Arquivos.MoveNext
            response.Flush()
        Wend
        RS_Arquivos.Close
      End If
      %>
    </table>
    </td>
  </tr>
<% End IF  %>
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