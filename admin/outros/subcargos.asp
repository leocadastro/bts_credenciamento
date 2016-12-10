<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
' * Dados Paginação
qtde					= Limpar_Texto(Request("qtd"))
ordem					= Limpar_Texto(Request("ord"))
acao_pagina				= Limpar_Texto(Request("acp"))
pagina					= Limpar_Texto(Request("pag"))
tabela					= Limpar_Texto(Request("tab"))

If tabela = "ramo" then
	campo_tabela = "Ramo"
ElseIf tabela = "atividade" then
	campo_tabela = "Atividade"
End If

empresa	= Limpar_Texto(Request("empresa"))
busca	= Limpar_Texto(Request("busca"))

if empresa <> "" then
	url_busca = url_busca & "&empresa=" & empresa
	Where_Busca = "and E.Razao_Social like '%" & empresa & "%'"
end if

if busca <> "" then
	url_busca = url_busca & "&busca=" & busca
	Where_Busca = Where_Busca & " and V.SubCargo_Outros like '%" & busca & "%'"
end if

strURL = "listar.asp?w=0"&url_busca

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
					"	E.Razao_Social as Empresa" &_
					"	,E.ID_Empresa " &_
					"	,V.ID_Visitante " &_
					"	,V.Nome_Completo " &_
					"	,V.Nome_Completo " &_
					"	,V.ID_SubCargo " &_
					"	,V.SubCargo_Outros " &_
					"	,V.Data_Cadastro " &_
					"From Visitantes as V "&_
					"Left Join Relacionamento_Cadastro as Rc ON V.ID_Visitante = Rc.ID_Visitante "&_
					"Left Join Empresas as E ON Rc.ID_Empresa = E.ID_Empresa "&_
					"Where V.SubCargo_Outros <> '' "&_
					Where_Busca &_
					"Order by V.SubCargo_Outros ASC, E.Razao_Social ASC"

	'response.write("<hr>" & SQL_Listar & "<hr>")

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
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">Subcargo Outros</span></td>
        <td width="100" align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;"><a href="/admin/outros/default.asp"><img src="/admin/images/bt_voltar_admin.gif" width="100" height="48" border="0" /></a></td>
      </tr>
    </table>
<div id="aviso" style="background-color:#FFFF00; width:100%; text-align:center;" class="t_arial fs11px bold c_preto"> <img src="/admin/images/ico_aviso.gif" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso"> <span id="aviso_conteudo">Aviso</span></div>

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
          	<form action="/admin/outros/subcargos.asp" method="post" name="busca">
                <span class="titulo_noticias_home" style="width: 500px; float: left;">
    
                    Empresa: <input type="text" name="empresa" id="empresa" class="admin_txtfield_login"> &nbsp; 
                    Subcargo: <input type="text" name="busca" id="busca" class="admin_txtfield_login">
                    
                    <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:right; width:15px; height:17px; float:right;" id="botao_dir"></div>
                    <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:right; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco cursor" onClick="document.busca.submit()">Buscar</div>
                    <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:right;"></div>
                </span>
            </form>
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
                <td width="275" class="borda_dir linha_16px" align="center"><b>Visitante</b></td>
                <td width="275" class="borda_dir linha_16px" align="center"><b>Empresa</b></td>
                <td width="200" class="borda_dir linha_16px" align="center"><b>Subcargo</b></td>
                <td width="50" align="center" class="borda_dir linha_16px"><b>Data</b></td>
                <td width="40" align="center" class="linha_16px"><b>Editar</b></td>
              </tr>
              <%
            Rs_listar.MoveFirst
            RS_Listar.AbsolutePage = PaginaAtual 
            While Not RS_Listar.EOF And Contador < RS_Listar.PageSize
            Contador = Contador + 1	
			
			SQL_Cadastro = 	"SELECT "&_
							"	ID_Relacionamento_Cadastro "&_
							"	,ID_Empresa "&_
						  	"FROM Relacionamento_Cadastro "&_
						  	"Where ID_Visitante = '" & RS_Listar("ID_Visitante") & "'"
			Set RS_Cadastro = Server.CreateObject("ADODB.Recordset")
			RS_Cadastro.Open SQL_Cadastro, Conexao, 3, 3
				If Not RS_Cadastro.Eof then
					ID_Cadastro = RS_Cadastro("ID_Relacionamento_Cadastro")
				End If
			RS_Cadastro.Close
			
			ID_Registro = "ID_Visitante"
	            %>
              <tr bgcolor="#FFFFFF" 
                onMouseOver="$(this).attr('bgcolor','#FFFF00'); " 
                onMouseOut="$(this).attr('bgcolor','#FFFFFF'); "  >
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'subcargos_editar.asp?tab=<%=tabela%>&id=<%=Rs_listar(ID_Registro)%>';" align="left"><%=RS_Listar(ID_Registro)%></td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'subcargos_editar.asp?tab=<%=tabela%>&id=<%=Rs_listar(ID_Registro)%>';" align="left"><%=RS_Listar("Nome_Completo")%></td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'subcargos_editar.asp?tab=<%=tabela%>&id=<%=Rs_listar(ID_Registro)%>';" align="left"><%=RS_Listar("Empresa")%></td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'subcargos_editar.asp?tab=<%=tabela%>&id=<%=Rs_listar(ID_Registro)%>';" align="left"><%=RS_Listar("SubCargo_Outros")%></td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'subcargos_editar.asp?tab=<%=tabela%>&id=<%=Rs_listar(ID_Registro)%>';" align="left"><%=FormatDateTIme(RS_Listar("Data_Cadastro"),2)%></td>
                <td class="borda_dir linha_16px cursor" align="center">
                	<a href="/admin/relatorios/relatorio_precredenciados_detalhes.asp?id=<%=ID_Cadastro%>">
                    	<img src="/admin/images/ico_preview_20_b.gif" width="15" height="15" alt="Detalhes do Cadastro" title="Detalhes do Cadastro" border="0"> 
                    </a>
                    <a href="subcargos_editar.asp?tab=<%=tabela%>&id=<%=RS_Listar(ID_Registro)%>">
                		<img src="/admin/images/ico_pg_prox.gif" width="15" height="15" alt="Atualizar" title="Atualizar" border="0">
                    </a>
                </td>
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