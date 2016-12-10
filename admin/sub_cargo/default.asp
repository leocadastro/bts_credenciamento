<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
' * Dados Paginação
qtde					= Limpar_Texto(Request("qtd"))
ordem					= Limpar_Texto(Request("ord"))
acao_pagina				= Limpar_Texto(Request("acp"))
pagina					= Limpar_Texto(Request("pag"))

campo_busca 			= Limpar_Texto(Request("campo_busca"))

If campo_busca <> "" Then
	URL_Busca = "&busca_cargo=" & campo_busca
	SQL_Busca = "Where SubCargo_PTB like '%" & campo_busca & "%' or SubCargo_ENG like '%" & campo_busca & "%' or SubCargo_ESP like '%" & campo_busca & "%' "
End If

strURL = "default.asp?w=0"&URL_Busca

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

	SQL_Cargos = 	"Select " &_
					"	ID_Cargo, " &_
					"	Cargo_PTB " &_
					"From Cargo " &_
					"Where Ativo = 1 " &_
					"	AND SubCargo = 1 " &_
					"Order by Cargo_PTB"
					
	Set RS_Cargos = Server.CreateObject("ADODB.Recordset")
	RS_Cargos.Open SQL_Cargos, Conexao

	SQL_Listar = 	"Select " &_
					"	S.ID_SubCargo " &_
					"	,C.Cargo_PTB as Cargo " &_
					"	,S.SubCargo_PTB " &_
					"	,S.SubCargo_ENG " &_
					"	,S.SubCargo_ESP " &_
					"	,S.Data_Cadastro " &_
					"	,S.Ativo " &_
					"From SubCargo as S " &_
					"Inner Join Cargo as C ON C.ID_Cargo = S.ID_Cargo " &_
					SQL_Busca &_
					"Order by ID_Subcargo DESC"

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
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">Sub-Cargos</span></td>
        <td width="270" align="center" valign="middle" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">
	        <a href="/admin/rel_edicoes_subcargo/"><img src="/admin/images/bt_link.png" height="48" border="0" style="float: left; margin-top: -5px;"/>
            <div id="relacionamento_link">Relacionar Sub-Cargos</div></a>
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
            <form id="cad" name="cad" method="post" action="metodos.asp">
            <input type="hidden" id="acao" name="acao" value="add_subcargo">
              <tr>
                <td height="30" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">Cadastrar</span></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Cargo:</td>
                <td class="titulo_noticias_home">
				<select id="id_cargo" name="id_cargo" class="admin_txtfield_login">
                <option value="-">-- Selecione --</option>
                <%
				If not RS_Cargos.BOF or not RS_Cargos.EOF Then
					While not RS_Cargos.EOF
					%><option value="<%=RS_Cargos("ID_Cargo")%>"><%=RS_Cargos("Cargo_PTB")%></option><%
						RS_Cargos.MoveNext
					Wend
					RS_Cargos.Close
				End If
				%>
                </select>
                </td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Sub-Cargo - PTB</td>
                <td class="titulo_noticias_home"><input name="subcargo_ptb" id="subcargo_ptb" type="text" class="admin_txtfield_login" size="50" /></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Sub-Cargo - ENG</td>
                <td class="titulo_noticias_home"><input name="subcargo_eng" id="subcargo_eng" type="text" class="admin_txtfield_login" size="50" /></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Sub-Cargo - ESP</td>
                <td class="titulo_noticias_home"><input name="subcargo_esp" id="subcargo_esp" type="text" class="admin_txtfield_login" size="50" /></td>
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
      
	<!--BUSCA-->
      <table width="600" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="14" height="12"><img src="/admin/images/caixa/top_esq.gif" width="14" height="12"></td>
          <td background="/admin/images/caixa/top_centro.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12"></td>
          <td width="14" height="12"><img src="/admin/images/caixa/top_dir.gif" width="14" height="12"></td>
        </tr>
        <tr>
          <td background="/admin/images/caixa/esq.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12"></td>
          <td>
            <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
                <form id="buscar" name="buscar" method="post" action="default.asp">
                  <tr>
                    <td height="30" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">Buscar</span></td>
                  </tr>
                  <tr>
                    <td height="30" class="titulo_menu_site_bts">Sub-Cargo</td>
                    <td class="titulo_noticias_home"><input name="busca_cargo" id="busca_cargo" type="text" class="admin_txtfield_login" size="30" /></td>
                  </tr>
                  <tr>
                    <td width="260" height="30" class="titulo_menu_site_tec">&nbsp;</td>
                    <td class="titulo_noticias_home"><div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:left;"></div>
                      <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco cursor" onClick="document.buscar.submit();">Buscar</div>
                      <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:right; width:15px; height:17px; float:left;"></div></td>
                  </tr>
                </form>
              </table>
          </td>
          <td background="/admin/images/caixa/dir.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12"></td>
        </tr>
        <tr>
          <td><img src="/admin/images/caixa/inf_esq.gif" width="14" height="12"></td>
          <td background="/admin/images/caixa/inf_centro.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12"></td>
          <td><img src="/admin/images/caixa/inf_dir.gif" width="14" height="12"></td>
        </tr>
      </table>

      <table width="900" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table>
	<!--BUSCA-->

      
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
                <td align="center" class="borda_dir linha_16px"><b>Cargo</b></td>
                <td class="borda_dir linha_16px" align="center"><b>Sub-Cargo</b></td>
                <td width="65" align="center" class="borda_dir linha_16px"><b>Disponível</b></td>
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
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'editar.asp?id=<%=Rs_listar("ID_SubCargo")%>';" align="left"><%=RS_Listar("ID_SubCargo")%></td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'editar.asp?id=<%=Rs_listar("ID_SubCargo")%>';" align="left"><%=Rs_listar("Cargo")%></td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'editar.asp?id=<%=Rs_listar("ID_SubCargo")%>';" align="left">
				<b>PTB:</b> <%=Rs_listar("SubCargo_PTB")%><br>
                <b>ENG:</b> <%=Rs_listar("SubCargo_ENG")%><br>
				<b>ESP:</b> <%=Rs_listar("SubCargo_ESP")%></td>
                <td class="borda_dir linha_16px cursor" align="center">
                <%
					ativo = Rs_listar("Ativo")
					If ativo = true OR ativo = 1 Then
						%><img src="/img/geral/icones/ok.gif" width="20" height="20" alt="ativo" title="ativo" onClick="document.location = 'metodos.asp?id=<%=Rs_listar("ID_SubCargo")%>&acao=desativar';"><%
					Else
						%><img src="/img/geral/icones/nok.gif" width="20" height="20" alt="desativado" title="desativado" onClick="document.location = 'metodos.asp?id=<%=Rs_listar("ID_SubCargo")%>&acao=ativar';"><%
                	End If
				%>
                </td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'editar.asp?id=<%=Rs_listar("ID_SubCargo")%>';" align="left"><%=FormatDateTIme(Rs_listar("Data_Cadastro"),2)%></td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'editar.asp?id=<%=Rs_listar("ID_SubCargo")%>';" align="center"><img src="/admin/images/ico_pg_prox.gif" width="15" height="15" alt="Ver Expositor" title="Ver Expositor" border="0"></td>
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