<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
' * Dados Paginação
qtde					= Limpar_Texto(Request("qtd"))
ordem					= Limpar_Texto(Request("ord"))
acao_pagina		= Limpar_Texto(Request("acp"))
pagina				= Limpar_Texto(Request("pag"))

campo_busca_fantasia	= Limpar_Texto(Request("campo_busca_fantasia"))
campo_busca_visitante = Limpar_Texto(Request("campo_busca_visitante"))
campo_busca_email		  = Limpar_Texto(Request("campo_busca_email"))

campo_busca_cnpj		= Limpar_Texto(Request("campo_busca_cnpj"))
campo_busca_cnpj		= Replace(campo_busca_cnpj,"/","")
campo_busca_cnpj		= Replace(campo_busca_cnpj,".","")
campo_busca_cnpj		= Replace(campo_busca_cnpj,"-","")

campo_busca_cpf			= Limpar_Texto(Request("campo_busca_cpf"))
campo_busca_cpf			= Replace(campo_busca_cpf,".","")
campo_busca_cpf			= Replace(campo_busca_cpf,"-","")

If campo_busca_fantasia <> "" Then
	URL_Busca = "&busca_cargo=" & campo_busca_fantasia
	WHERE_Busca = "Where E.Nome_Fantasia like '%" & campo_busca_fantasia & "%' "
End If

If campo_busca_cnpj <> "" Then
	URL_Busca = "&busca_cargo=" & campo_busca_cnpj
	
	If WHERE_Busca <> "" then
		WHERE_Busca = WHERE_Busca & " and E.CNPJ = '" & campo_busca_cnpj & "'"
	Else
		WHERE_Busca = "Where E.CNPJ = '" & campo_busca_cnpj & "'"
	End If
End If

If campo_busca_visitante <> "" Then
	URL_Busca = URL_Busca & "&busca_cargo=" & campo_busca_visitante
	
	If WHERE_Busca <> "" then
		WHERE_Busca = WHERE_Busca & " and V.Nome_Credencial like '%" & campo_busca_visitante & "%'"
	Else
		WHERE_Busca = "Where V.Nome_Credencial like '%" & campo_busca_visitante & "%'"
	End If
End If

If campo_busca_cpf <> "" Then
	URL_Busca = URL_Busca & "&busca_cargo=" & campo_busca_cpf
	
	If WHERE_Busca <> "" then
		WHERE_Busca = WHERE_Busca & " and V.CPF = '" & campo_busca_cpf & "'"
	Else
		WHERE_Busca = " Where V.CPF = '" & campo_busca_cpf & "'"
	End If
End If

If campo_busca_email <> "" Then
	URL_Busca = URL_Busca & "&busca_cargo=" & campo_busca_email
	
	If WHERE_Busca <> "" then
		WHERE_Busca = WHERE_Busca & " and V.email = '" & campo_busca_email & "'"
	Else
		WHERE_Busca = " Where V.email = '" & campo_busca_email & "'"
	End If
End If

strURL = "relatorio_precredenciados.asp?w=0"

If Len(qtde) > 0 AND isNumeric(qtde) Then 
	qtde_itens = qtde
	strURL = strURL & "&qtd=" & qtde
Else
	qtde_itens = 40
	qtde = 40
End If

' * Request BUSCA
id_edicao         = Limpar_Texto(Request("id_edicao"))

condicoes = ""
If id_edicao <> "" Then
  strURL = strURL & "&id_edicao=" & id_edicao
	If WHERE_Busca <> "" then
		condicoes = " Where RC.ID_Edicao = " & id_edicao & " "
		WHERE_Busca = WHERE_Busca & " and RC.ID_Edicao = " & id_edicao & " "
	Else
		condicoes = " Where RC.ID_Edicao = " & id_edicao & " "
  		WHERE_Busca = "Where RC.ID_Edicao = " & id_edicao & " "
	End If
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

	
  SQL_Contar =  "Select Count(RC.ID_Relacionamento_Cadastro) as ID " &_
                "From Relacionamento_Cadastro as RC " &_
                "Inner Join Tipo_Credenciamento as TC ON TC.ID_Tipo_Credenciamento = RC.ID_Tipo_Credenciamento " &_
                "Inner Join Formularios as F ON F.ID_Formulario = TC.ID_Formulario " &_
                "Inner Join Eventos_Edicoes as EE ON EE.ID_Edicao = RC.ID_Edicao " &_
                "Inner Join Eventos as EV ON EV.ID_Evento = EE.ID_Evento " &_
                "Inner Join Visitantes as V ON V.ID_Visitante = RC.ID_Visitante " &_
                "Left  Join Empresas as E ON E.ID_Empresa = RC.ID_Empresa " &_
                " " & condicoes & " "
  'response.write("<hr>" & SQL_Contar & "<hr>")

  Set RS_Contar = Server.CreateObject("ADODB.Recordset")
  RS_Contar.Open SQL_Contar, Conexao, 3, 3


  'Quantidade de cadastros
  QtdeCadastros = RS_Contar("ID")

  if id_edicao = "1" Then ' CARDS 2012 
    BdAntigo = " + 288 Banco Antigo"
  End If

  SQL_Listar = 	"Select " &_
                "     RC.ID_Relacionamento_Cadastro as ID " &_
                "     ,F.Nome as Formulario " &_
                "     ,V.Nome_Credencial as Visitante " &_
                "     ,E.Nome_Fantasia as Empresa " &_
                "     ,EV.Nome_PTB as Evento " &_
                "     ,EE.Ano as Ano " &_
                "From Relacionamento_Cadastro as RC " &_
                "Inner Join Tipo_Credenciamento as TC ON TC.ID_Tipo_Credenciamento = RC.ID_Tipo_Credenciamento " &_
                "Inner Join Formularios as F ON F.ID_Formulario = TC.ID_Formulario " &_
                "Inner Join Eventos_Edicoes as EE ON EE.ID_Edicao = RC.ID_Edicao " &_
                "Inner Join Eventos as EV ON EV.ID_Evento = EE.ID_Evento " &_
                "Inner Join Visitantes as V ON V.ID_Visitante = RC.ID_Visitante " &_
                "Left  Join Empresas as E ON E.ID_Empresa = RC.ID_Empresa " &_
                " " & WHERE_Busca & " " &_
                "Order by ID DESC "
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
	
	$("#campo_busca_cnpj").mask("99.999.999/9999-99",{placeholder:"_"});
	$("#campo_busca_cpf").mask("999.999.999-99",{placeholder:"_"});
	$("#campo_busca_cnpj").val('');
	$('#aviso').hide();
	

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
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">Listagem de Pré-Credenciados</span></td>
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
            <form id="buscar" name="buscar" method="post" action="relatorio_precredenciados.asp">
              <tr>
                <td height="30" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">Buscar Pré-Credenciado</span></td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_bts"> Edi&ccedil;&atilde;o</td>
                <td class="titulo_noticias_home">
                <select id="id_edicao" name="id_edicao" class="admin_txtfield_login">
                <option value="">-- Selecione --</option>
                <%
                  If not RS_Eventos.BOF or not RS_Eventos.EOF Then
                    While not RS_Eventos.EOF
                    %><option value="<%=RS_Eventos("ID_Edicao")%>"><%=RS_Eventos("ano")%> - <%=RS_Eventos("Evento")%></option><%
                      RS_Eventos.MoveNext
                    Wend
                    RS_Eventos.Close
                  End If
                  %>
                </select>
                </td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">CNPJ</td>
                <td class="titulo_noticias_home"><input name="campo_busca_cnpj" id="campo_busca_cnpj" type="text" class="admin_txtfield_login" size="30" /></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Fantasia</td>
                <td class="titulo_noticias_home"><input name="campo_busca_fantasia" id="campo_busca_fantasia" type="text" class="admin_txtfield_login" size="30" /></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">CPF</td>
                <td class="titulo_noticias_home"><input name="campo_busca_cpf" id="campo_busca_cpf" type="text" class="admin_txtfield_login" size="30" /></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Visitante</td>
                <td class="titulo_noticias_home"><input name="campo_busca_visitante" id="campo_busca_visitante" type="text" class="admin_txtfield_login" size="30" /></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">E-mail</td>
                <td class="titulo_noticias_home"><input name="campo_busca_email" id="campo_busca_email" type="text" class="admin_txtfield_login" size="30" /></td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_tec">&nbsp;</td>
                <td class="titulo_noticias_home">
                  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:left;"></div>
                  <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco cursor" onClick="document.buscar.submit();">Buscar</div>
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
      </table><br/>
      
      
      <span class="titulo_menu_site_tec">TOTAL DE PRÉ-CREDENCIADOS <b><%=QtdeCadastros%> <%=BdAntigo%></b></span>
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
                <td class="borda_dir linha_16px" align="center"><b>Credenciamento</b></td>
                <td class="borda_dir linha_16px" align="center"><b>Fantasia</b></td>
                <td class="borda_dir linha_16px" align="center"><b>Visitante</b></td>
                <td class="borda_dir linha_16px" align="center"><b>Evento</b></td>
                <td width="40" align="center" class="linha_16px"><b>Detalhes</b></td>
              </tr>
              <%
            Rs_listar.MoveFirst
            RS_Listar.AbsolutePage = PaginaAtual 
            While Not RS_Listar.EOF And Contador < RS_Listar.PageSize
            Contador = Contador + 1	 
            
            ID = RS_Listar("ID")
            If isNull(RS_Listar("Formulario")) = true Then 
              Formulario = "-"
            Else
              Formulario = RS_Listar("Formulario")
            End If

            If isNull(Rs_listar("Empresa")) = true Then
              Empresa = "-"
            Else
              Empresa = Rs_listar("Empresa")
            End If

            If isNull(Rs_listar("Visitante")) = true Then
              Visitante = "-"
            Else
              Visitante = Rs_listar("Visitante")
            End If
            Evento = Rs_listar("Ano") & " " & Rs_listar("Evento")
            			
            %>
              <tr bgcolor="#FFFFFF" 
                onMouseOver="$(this).attr('bgcolor','#FFFF00'); " 
                onMouseOut="$(this).attr('bgcolor','#FFFFFF'); "  >
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'relatorio_precredenciados_detalhes.asp?id=<%=ID%>';" align="left"><%=ID%></td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'relatorio_precredenciados_detalhes.asp?id=<%=ID%>';" align="left"><%=Formulario%></td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'relatorio_precredenciados_detalhes.asp?id=<%=ID%>';" align="left"><%=Empresa%></td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'relatorio_precredenciados_detalhes.asp?id=<%=ID%>';" align="left"><%=Visitante%></td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'relatorio_precredenciados_detalhes.asp?id=<%=ID%>';" align="left"><%=Evento%></td>
                <td class="borda_dir linha_16px cursor" onClick="document.location = 'relatorio_precredenciados_detalhes.asp?id=<%=ID%>';" align="center"><img src="/admin/images/ico_busca_24px.gif" width="15" height="15" alt="Ver Expositor" title="Ver Expositor" border="0"></td>
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