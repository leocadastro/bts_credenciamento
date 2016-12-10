<!--#include virtual="/admin/inc/logado.asp"-->
<html>
<head>
<meta http-equiv="Content-type" content="text/html; charset=iso-8859-1" />
<title>Administração Cred. 2012</title>
<link href="/admin/css/bts.css" rel="stylesheet" type="text/css" />
<link href="/admin/css/admin.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
</head>

<%
' Níveis
' 1 - Admin
' 2 - Marketing
' 3 - Editor

Dim menu(29)
'* Na ordem de exibição 
'  ( 0 - 1 ) 
'  ( 2 - 3 ) ...
'MODELO - Array (titulo, icone, link, permissao)
menu(0) = Array("Eventos", "spacer", "document.location='/admin/eventos/';", "1,2")
menu(1) = Array("Edições dos Eventos", "spacer", "document.location='/admin/edicoes/';", "1,2")
menu(2) = Array("Configurar Edições", "spacer", "document.location='/admin/edicoes_visual/';", "1,2")
menu(3) = Array("Tipos de Credenciamento", "spacer", "document.location='/admin/tipo_credenciamento/';", "1,2")
menu(4) = Array("Relacionar - Edições > Tipos", "spacer", "document.location='/admin/rel_tipos_edicoes/';", "1,2")
menu(5) = Array("Tradução Páginas", "spacer", "document.location='/admin/paginas_web/';", "1,2")
menu(6) = Array("Cargos", "spacer", "document.location='/admin/cargo/';", "1,2")
menu(7) = Array("Sub-Cargos", "spacer", "document.location='/admin/sub_cargo/';", "1,2")
menu(8) = Array("Departamentos", "spacer", "document.location='/admin/depto/';", "1,2")
menu(9) = Array("Ramos V2", "spacer", "document.location='/admin/ramos_v2/';", "1,2")
menu(10) = Array("Atividade Econômica", "spacer", "document.location='/admin/atividade/';", "1,2")
menu(11) = Array("Área de Interesse", "spacer", "document.location='/admin/area_interesse/';", "1,2")
menu(12) = Array("Área de Atuação", "spacer", "document.location='/admin/area_atuacao/';", "1,2")
menu(13) = Array("Interesse na Feira", "spacer", "document.location='/admin/interesse_feira/';", "1,2")
menu(14) = Array("Quantidade de funcionários", "spacer", "document.location='/admin/funcionarios/';", "1,2")
menu(15) = Array("Relatórios", "spacer", "document.location='/admin/relatorios/';", "1,2,3")
menu(16) = Array("Relacionar - Edições > Ramos V2", "spacer", "document.location='/admin/rel_edicoes_ramos_v2/';", "1,2")
menu(17) = Array("Relacionar - Edições > Atividade", "spacer", "document.location='/admin/rel_edicoes_atividade/';", "1,2")
menu(18) = Array("Relacionar - Edições > Áreas de Interesse", "spacer", "document.location='/admin/rel_edicoes_interesse/';", "1,2")
menu(19) = Array("Relacionar - Edições > Áreas de Atuação", "spacer", "document.location='/admin/rel_edicoes_atuacao/';", "1,2")
menu(20) = Array("Relacionar - Edições > Interesse na Feira", "spacer", "document.location='/admin/rel_edicoes_interessefeira/';", "1,2")
menu(21) = Array("Relacionar - Edições > Cargos", "spacer", "document.location='/admin/rel_edicoes_cargo/';", "1,2")
menu(22) = Array("Relacionar - Edições > Sub-Cargos", "spacer", "document.location='/admin/rel_edicoes_subcargo/';", "1,2")
menu(23) = Array("Exportar Pré-Credenciados", "spacer", "document.location='/admin/exportar_xml/';", "1,2,4") '4'
menu(24) = Array("Produtos", "spacer", "document.location='/admin/produtos/';", "1") '4'
menu(25) = Array("Administração de OUTROS", "spacer", "document.location='/admin/outros/';", "1") '4'
menu(26) = Array("Perguntas", "spacer", "document.location='/admin/perguntas/';", "1,2") '4'
menu(27) = Array("Relacionar - Edições > Perguntas", "spacer", "document.location='/admin/rel_edicoes_perguntas/';", "1,2") '4'
menu(28) = Array("Gerador de Links Diretos pra Feira", "spacer", "document.location='/admin/links_feira/';", "1,2") '4'
menu(29) = Array("F.A.Q", "spacer", "document.location='/admin/faq/';", "1,2") '4'
%>

<script language="javascript">
$(document).ready(function(){
	$('#aviso').hide();
	<% 
	'msg = Request("msg")
	If msg = "" AND Session("admin_msg") <> "" Then msg = Session("admin_msg")
	Select Case msg
		Case "pag_proibida"
			Session("admin_msg") = ""
		%>
		$('#aviso_conteudo').html('Seu usuário não tem permissão para acessar a página solicitada !');
		$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000);
		<%
	End Select
	%>
});

// INICIO - Configuração de navegar por TECLAS
var links = new Array();
<%
For i = Lbound(menu) to Ubound(menu)
	valor = Replace(menu(i)(2),"document.location=","")
	valor = Replace(valor, "window.open","")
	valor = Replace(valor, "(","")
	valor = Replace(valor, ")","")
	%>links[<%=i+1%>] = <%=valor%><%
Next
%>
var tecla = '';
function showKeyPress(evt)
{
	tecla += String.fromCharCode(evt.charCode);
	if (tecla.length >= 2) {
		verificar_tecla();
	} else {
		setTimeout(function() {
			verificar_tecla();
		}, 400);
	}
}
function verificar_tecla() {
	for (i = 1; i < links.length; i++) {
		if (tecla.toString() == i.toString()) {
			document.location = links[i];
		} 
	}
	tecla = '';	
}
// FIM - Configuração de navegar por TECLAS
</script>

<body onkeypress="showKeyPress(event);">

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
    <td align="center" bgcolor="#FFFFFF"><!-- ************** CONTEUDO ************** -->
      <table width="560" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="50" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center">Menu Principal <br>
            <span class="fs10px c_azul1">Dica: navegue digitando o n&uacute;mero correspondente</span></td>
        </tr>
      </table>
      <div id="aviso" style="background-color:#FFFF00; width:100%; text-align:center;" class="t_arial fs11px bold c_preto"> <img src="/admin/images/ico_aviso.gif" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso"> <span id="aviso_conteudo">Aviso</span> </div>
      <br>
      <% 'Loop Menu %>
      <table width="560" border="0" cellspacing="0" cellpadding="0">
        <% 
	colunas = 0
	For i = LBound(menu) to Ubound(menu)
		exibir = false
		'Array (titulo, icone, link, permissao)
		'Array (0     , 1    , 2   , 3        )
		'response.Write(menu(i)(0) & " / " & menu(i)(1) & " / " & menu(i)(2) & " / " & menu(i)(3) & "<br>")
		permissao_item = Split(menu(i)(3), ",")
		For p = LBound(permissao_item) to Ubound(permissao_item)
			If Cstr(Session("admin_id_perfil")) = Cstr(permissao_item(p)) Then
				exibir = true
			End If
		Next
		'response.write(colunas & "-")
		If colunas = 0 Then 
			'response.write(colunas & "-")
			colunas = colunas + 1
			%>
        <tr valign="top">
          <%
		End If

		If exibir = true Then 
			If colunas = 1 or colunas = 3 Then
			'response.write(colunas & "-")
			colunas = colunas + 1
			%>
          <td><table width="271" border="0" cellspacing="0" cellpadding="0" background="/admin/images/bts/fundo_bts_menu.gif" class="cursor" onClick="<%=menu(i)(2)%>">
            <tr>
            
              <% If menu(i)(1) <> "spacer" Then%>
              <td width="74"><img src="/admin/images/bts/<%=menu(i)(1)%>.gif" width="58" height="48" hspace="8"></td>
			  <% Else %>
              <td width="74" class="c_vermelho fs22px t_arial bold" align="center"><%=i + 1%></td>
              <% End If %>
              <td height="54" class="bt_menu_titulo_home fs12px" style="padding-right:4px;"><%=menu(i)(0)%></td>
            </tr>
          </table></td>
          <%
			End If
		End If
		If colunas = 2 Then
			'response.write(colunas & "-")
			colunas = colunas + 1
		%>
          <td width="20" height="70">&nbsp;</td>
          <%
		ElseIf colunas = 4 Then 
			colunas = 0
		%>
        </tr>
        <%
		End If
    Next
	%>
      </table>
      <!-- ************** CONTEUDO ************** --></td>
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