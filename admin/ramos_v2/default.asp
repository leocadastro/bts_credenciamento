<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
' * Dados Pagina��o
qtde					= Limpar_Texto(Request("qtd"))
ordem					= Limpar_Texto(Request("ord"))
acao_pagina				= Limpar_Texto(Request("acp"))
pagina					= Limpar_Texto(Request("pag"))

strURL = "default.asp?w=0"

If Len(qtde) > 0 AND isNumeric(qtde) Then 
	qtde_itens = qtde
	strURL = strURL & "&qtd=" & qtde
Else
	qtde_itens = 20
	qtde = 20
End If
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")

	SQL_Listar = 	"Select " &_
					"	ID_Ramo_Atividade " &_
					"	,Ramo_Atv_PTB " &_
					"	,Ramo_Atv_ENG " &_
					"	,Ramo_Atv_ESP " &_
					"	,Complemento " &_
					"	,Ativo " &_
					"	,Data_Cadastro " &_
					"From RamoeAtividade_V2 " &_
					"Order by ID_Ramo_Atividade DESC"

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
<title>Administra��o Cred. 2012</title>
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
			$('#aviso_conteudo').html('P�gina n�o permitida !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 10000);
			<%
		Case "add_ok"
			%>
			$('#aviso_conteudo').html('Item adicionado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 10000).fadeOut('slow');
			<%
		Case "erro_nao_encontrado"
			%>
			$('#aviso_conteudo').html('Erro - N�o foi encontrado nenhum registro !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 10000).fadeOut('slow');
			<%
		Case "add_erro_existe"
			%>
			$('#aviso_conteudo').html('Erro - Item informado j� existe !');
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
		// Se n�o for obrigat�rio
		switch (this.id) {
			default:
				erros += verificar(this.id, false);
				break;
		}
	});
	$('input:enabled').each(function(i) {
		// Se n�o for obrigat�rio
		switch (this.id) {
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
        <td style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">Ramos e Atividades V2</span></td>
        <td width="250" align="center" valign="middle" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">
	        <a href="/admin/rel_edicoes_ramos_v2/"><img src="/admin/images/bt_link.png" height="48" border="0" style="float: left; margin-top: -5px;"/>
            <div id="relacionamento_link">Relacionar Ramos V2</div></a>
        </td>
        <td width="100" align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">
        	<a href="/admin/menu.asp"><img src="/admin/images/bt_voltar_admin.gif" width="100" height="48" border="0" /></a>
        </td>
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
            <input type="hidden" id="acao" name="acao" value="add_Ramo">
              <tr>
                <td height="30" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">Cadastrar</span></td>
                </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Ramo - PTB</td>
                <td class="titulo_noticias_home"><input name="ramo_ptb" id="ramo_ptb" type="text" class="admin_txtfield_login" size="30" /></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Ramo - ENG</td>
                <td class="titulo_noticias_home"><input name="ramo_eng" id="ramo_eng" type="text" class="admin_txtfield_login" size="30" /></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Ramo - ESP</td>
                <td class="titulo_noticias_home"><input name="ramo_esp" id="ramo_esp" type="text" class="admin_txtfield_login" size="30" /></td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_bts">Possui Complemento</td>
                <td class="titulo_noticias_home">
                  <select id="complemento" name="complemento" class="admin_txtfield_login">
                    <option value="0">N�o</option>
                    <option value="1">Sim</option>
                  </select>
                  </td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_bts">Dispon&iacute;vel</td>
                <td class="titulo_noticias_home">
                  <select id="ativo" name="ativo" class="admin_txtfield_login">
                    <option value="0">N�o</option>
                    <option value="1">Sim</option>
                  </select>
                  </td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_tec">&nbsp;</td>
                <td class="titulo_noticias_home">
                  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:left;"></div>
                  <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco cursor" onClick="Enviar();">Cadastrar </div>
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
            <table width="100%" border="0" cellpadding="2" cellspacing="1" class="fs11px t_arial">
              <tr>
                <td width="50" align="center" class="borda_dir linha_16px"><b>ID</b></td>
                <td class="borda_dir linha_16px" align="center"><b>Portugu&ecirc;s</b></td>
                <td class="borda_dir linha_16px" align="center"><b>Ingl&ecirc;s</b></td>
                <td class="borda_dir linha_16px" align="center"><b>Espanhol</b></td>
                <td width="65" align="center" class="borda_dir linha_16px"><b>Possui Complemento</b></td>
                <td width="65" align="center" class="borda_dir linha_16px"><b>Dispon�vel</b></td>
                <td width="100" align="center" class="borda_dir linha_16px"><b>Data</b></td>
                <td width="40" align="center" class="linha_16px"><b>Editar</b></td>
              </tr>
              <%
            Rs_listar.MoveFirst
            RS_Listar.AbsolutePage = PaginaAtual 
            While Not RS_Listar.EOF And Contador < RS_Listar.PageSize
            Contador = Contador + 1				
            %>
              <tr bgcolor="#FFFFFF" 
                onMouseOver="$(this).attr('bgcolor','#FFFF00'); " 
                onMouseOut="$(this).attr('bgcolor','#FFFFFF'); "  >
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'editar.asp?id=<%=Rs_listar("ID_Ramo_Atividade")%>';" align="left"><%=RS_Listar("ID_Ramo_Atividade")%></td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'editar.asp?id=<%=Rs_listar("ID_Ramo_Atividade")%>';" align="left"><%=Rs_listar("Ramo_Atv_PTB")%></td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'editar.asp?id=<%=Rs_listar("ID_Ramo_Atividade")%>';" align="left"><%=Rs_listar("Ramo_Atv_ENG")%></td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'editar.asp?id=<%=Rs_listar("ID_Ramo_Atividade")%>';" align="left"><%=Rs_listar("Ramo_Atv_ESP")%></td>
                <td class="borda_dir linha_16px " align="center">
                <%
					complemento = Rs_listar("complemento")
					If complemento = true OR complemento = 1 Then
						%><img src="/img/geral/icones/ok.gif" width="20" height="20" title="ativo"><%
					Else
						%><img src="/img/geral/icones/nok.gif" width="20" height="20" title="desativado"><%
                	End If
				%>
                </td>
                <td class="borda_dir linha_16px cursor" align="center">
                <%
					ativo = Rs_listar("Ativo")
					If ativo = true OR ativo = 1 Then
						%><img src="/img/geral/icones/ok.gif" width="20" height="20" alt="ativo" title="ativo" onClick="document.location = 'metodos.asp?id=<%=Rs_listar("ID_Ramo_Atividade")%>&acao=desativar';"><%
					Else
						%><img src="/img/geral/icones/nok.gif" width="20" height="20" alt="desativado" title="desativado" onClick="document.location = 'metodos.asp?id=<%=Rs_listar("ID_Ramo_Atividade")%>&acao=ativar';"><%
                	End If
				%>
                </td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'editar.asp?id=<%=Rs_listar("ID_Ramo_Atividade")%>';" align="left"><%=FormatDateTIme(Rs_listar("Data_Cadastro"),2)%></td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'editar.asp?id=<%=Rs_listar("ID_Ramo_Atividade")%>';" align="center"><img src="/admin/images/ico_pg_prox.gif" width="15" height="15" alt="Ver Expositor" title="Ver Expositor" border="0"></td>
              </tr>
              <%
                RS_Listar.MoveNext()
            Wend
            RS_Listar.Close
			%>
            </table>
            <%
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