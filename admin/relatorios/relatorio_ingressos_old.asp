<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
'if session("admin_id_usuario") <> "21" then
'	response.Redirect("/admin/relatorios/default.asp")
'end if
' * Dados Paginação
Dim ano_pedido, campo_ordem, tipo_ordem

qtde			= Limpar_Texto(Request("qtd"))
ordem			= Limpar_Texto(Request("ord"))
acao_pagina		= Limpar_Texto(Request("acp"))
pagina			= Limpar_Texto(Request("pag"))
id_status 		= Limpar_Texto(Request("id_status"))
pedido	 		= Limpar_Texto(Request("pedido"))
campo_busca		= Limpar_Texto(Request("campo_busca"))
ID_Evento 		= Limpar_Texto(Request("ID_Evento"))

ano_pedido = Limpar_Texto(Request("ano_pedido"))
campo_ordem = Limpar_Texto(Request("campo_ordem"))
tipo_ordem = Limpar_Texto(Request("tipo_ordem"))

If Len(qtde) > 0 AND isNumeric(qtde) Then 
	qtde_itens = qtde
	strURL = strURL & "&qtd=" & qtde
Else
	qtde_itens = 20
	qtde = 20
End If

strURL = "relatorio_ingressos.asp?w=0&id_evento="& id_evento & "&campo_ordem=" & campo_ordem & "&tipo_ordem=" & tipo_ordem & "&id_status=" & id_status&"&ano_pedido=" & ano_pedido & "&campo_busca=" & campo_busca & "&pedido=" & pedido & "" & strURL

If id_status <> "" Then
	SQL_Where = SQL_Where & "	And Pe.Status_Pedido = " & id_status & " "
End If

If pedido <> "" Then
	Select Case campo_busca
		Case "numeropedido"
			campo = "Pe.Numero_Pedido"
		Case "nomecomprador"
			campo = "Vc.Nome_Completo"
		Case "cpfcomprador"
			campo = "Vc.CPF"
	End Select
	SQL_Where = SQL_Where & "	And " & campo & " like '%" & pedido & "%' "
End If

If ano_pedido <> "" And IsNumeric(ano_pedido) Then
    SQL_Where = SQL_Where & "	And YEAR(Pe.Data_Pedido) = " & ano_pedido
End If

'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")

	SQL_Listar =   	"Select " &_
					"	Count(Pc.ID_Pedido) As Quantidade " &_
					"	,Pe.ID_Pedido " &_
					"	,Pe.ID_Edicao " &_
					"	,Pe.ID_Idioma " &_
					"	,Pe.Numero_Pedido " &_
					"	,Pe.ID_Rel_Cadastro " &_
					"	,Pe.ID_Visitante As ID_Comprador " &_
					"	,Vc.Nome_Completo As Nome_Comprador " &_
					"	,Vc.CPF As CPF_Comprador " &_
					"	,Vc.Senha " &_
					"	,Pe.Valor_Pedido " &_
					"	,Pe.Data_Pedido " &_
					"	,Pe.Status_Pedido As ID_Status " &_
					"	,St.Status_PTB As Status_Pedido " &_
					" From Pedidos_Carrinho As Pc WITH(NOLOCK) " &_
					" Inner Join Pedidos As Pe WITH(NOLOCK) " &_
					"	On Pc.ID_Pedido = Pe.ID_Pedido " &_
					" Inner Join Pedidos_Status As St WITH(NOLOCK) " &_
					"	On St.ID_Pedido_Status = Pe.Status_Pedido " &_
					" Inner Join Visitantes As Vc WITH(NOLOCK) " &_
					"	On Vc.ID_Visitante = Pe.ID_Visitante " &_
					" Inner Join Eventos_Edicoes As eve_edi WITH(NOLOCK) " &_
					"	On eve_edi.ID_Edicao = Pe.ID_Edicao " &_
					"	AND eve_edi.ID_Evento = '" & ID_Evento & "'" &_
					" Where Pe.Valor_Pedido > 1" &_
					SQL_Where &_
					" Group By " &_
					"	Pc.ID_Pedido " &_
					"	,Pe.ID_Pedido " &_
					"	,Pe.ID_Edicao " &_
					"	,Pe.ID_Idioma " &_
					"	,Pe.Data_Pedido " &_
					"	,Pe.ID_Rel_Cadastro " &_
					"	,Pe.ID_Visitante " &_
					"	,Vc.Nome_Completo " &_
					"	,Vc.CPF " &_
					"	,Vc.Senha " &_
					"	,Pe.Valor_Pedido " &_
					"	,Pe.Numero_Pedido " &_
					"	,Pe.Status_Pedido " &_
					"	,St.Status_PTB "

	If campo_ordem <> "" Then
		Select Case campo_ordem
			Case "datapedido"
				ordem = "Pe.Data_Pedido"
			Case "nomecomprador"
				ordem = "Vc.Nome_Completo"
			Case "numeropedido"
				ordem = "Pe.Numero_Pedido"
		End Select
		
		Select Case tipo_ordem
			Case "crescente"
				tipo = "ASC"
			Case "decrescente"
				tipo = "DESC"
		End Select
		
		SQL_Listar = SQL_Listar & " ORDER BY " & ordem & " " & tipo
	End If
	
  'Response.Write(SQL_Listar)
  'response.end
	Set RS_Listar = Server.CreateObject("ADODB.Recordset")
	RS_Listar.CursorLocation = 3
	RS_Listar.CacheSize = qtde_itens
	RS_Listar.PageSize = qtde_itens
	RS_Listar.Open SQL_Listar, Conexao
  
	If not RS_Listar.BOF and not RS_Listar.EOF Then
		TotalPaginas = RS_Listar.PageCount	
	Else 
		TotalPaginas = 0 
	End IF
  
  
	SQL_Soma =   	"Select Distinct " &_
					"	Pc.Id_Pedido " &_
					"	,Pe.Numero_Pedido " &_ 
					"	,Pe.Status_Pedido  " &_
					"	,Count(Pc.ID_Pedido) As Quantidade " &_
					"From Pedidos_Carrinho As Pc WITH(NOLOCK)  " &_
					"Inner Join Pedidos As Pe WITH(NOLOCK)  " &_
					"	On Pc.ID_Pedido = Pe.ID_Pedido  " &_
					"Inner Join Pedidos_Status As St WITH(NOLOCK)  " &_
					"	On St.ID_Pedido_Status = Pe.Status_Pedido  " &_
					"Inner Join Visitantes As Vc WITH(NOLOCK)  " &_
					"	On Vc.ID_Visitante = Pe.ID_Visitante  " &_
					" Inner Join Eventos_Edicoes As eve_edi WITH(NOLOCK) " &_
					"	On eve_edi.ID_Edicao = Pe.ID_Edicao " &_
					"	AND eve_edi.ID_Evento = " & ID_Evento &_
					" Where Pe.Valor_Pedido > 1" &_
					SQL_Where &_
					"Group By " &_
					"	Pc.Id_Pedido " &_
					"	,Pe.Numero_Pedido " &_ 
					"	,Pe.Status_Pedido "
  'Response.Write(SQL_Soma)
  'Response.end
  Set RS_Soma = Server.CreateObject("ADODB.Recordset")
  RS_Soma.Open SQL_Soma, Conexao
    
  
  SQL_Status =	"Select ID_Pedido_Status, Status_PTB, Status_ENG, Status_ESP " &_
				"From Pedidos_Status WITH(NOLOCK)"
  Set RS_Status = Server.CreateObject("ADODB.Recordset")
  RS_Status.Open SQL_Status, Conexao
  
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

If ID_Evento = 5 Then
	Titulo_Pagina = "Listagem de Ingressos - ABF"
ElseIf ID_Evento = 16 Then
	Titulo_Pagina = "Listagem de Ingressos - ABF Nordeste"
End If

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


function cancelar_pedido(id_pedido){

window.location = "cancelar_pedido.asp?id_pedido=" + id_pedido + "&id_status=<%=id_status%>&pag=<%=pagina%>" ;

}
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
function Page_Load()
{
	document.buscar.campo_ordem.value = "<%=campo_ordem%>";
	document.buscar.tipo_ordem.value = "<%=tipo_ordem%>";
	if ("<%=campo_busca%>" != "")
		document.buscar.campo_busca.value = "<%=campo_busca%>";
}
function Exportar_Excel()
{
	var acaoAnterior = document.buscar.action;
	document.buscar.action = "relatorio_ingressos_exportar.asp";
	document.buscar.submit();
	
	document.buscar.action = acaoAnterior;
}
</script>

<body onload="Page_Load()">
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
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22"><%=Titulo_Pagina%></span></td>
        <td width="100" align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;"><a href="default.asp"><img src="/admin/images/bt_voltar_admin.gif" width="100" height="48" border="0" /></a></td>
      </tr>
    </table>
<div id="aviso" style="background-color:#FFFF00; width:100%; text-align:center;" class="t_arial fs11px bold c_preto"> <img src="/admin/images/ico_aviso.gif" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso"> <span id="aviso_conteudo">Aviso</span></div>

	<div style="float: left; margin-left: 30px;">
      <table width="350" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="14" height="12"><img src="/admin/images/caixa/top_esq.gif" width="14" height="12" /></td>
          <td background="/admin/images/caixa/top_centro.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td width="14" height="12"><img src="/admin/images/caixa/top_dir.gif" width="14" height="12" /></td>
        </tr>
        <tr>
          <td background="/admin/images/caixa/esq.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td style="height:152px">
          
          <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td width="200" height="30" class="titulo_menu_site_bts" style="line-height: 250%">
                	<%
					Soma_Total 		= 0
					Soma_Novo 		= 0
					Soma_Pago 		= 0
					Soma_Cancelado	= 0
							
					While Not RS_Soma.Eof
					
						Soma_Total = Soma_Total + Cint(RS_Soma("Quantidade"))
						
						If Cint(RS_Soma("Status_Pedido")) = 1 Then
						
							Soma_Novo = Soma_Novo + Cint(RS_Soma("Quantidade"))
							
						ElseIf Cint(RS_Soma("Status_Pedido")) = 3 Then
						
							Soma_Pago = Soma_Pago + Cint(RS_Soma("Quantidade"))
							
						ElseIf Cint(RS_Soma("Status_Pedido")) = 4 Then
						
							Soma_Cancelado = Soma_Cancelado + Cint(RS_Soma("Quantidade"))
							
						End If
					
					RS_Soma.MoveNext
					Wend
					
					RS_Soma.Close
					%>
                    
                    <font size="+1">Total de Ingressos:</font> <br>
                    Novos Ingressos: <br>
                    Ingressos Pagos: <br>
                    Ingressos Cancelados: 
                </td>
                <td class="titulo_noticias_home" style="line-height: 250%">
                	<font size="+1"><%=Soma_Total%></font> <br>
                    <%=Soma_Novo%> <br>
                    <%=Soma_Pago%> <br>
                    <%=Soma_Cancelado%>
                </td>
              </tr>
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
      </div>




      <table width="520" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="14" height="12"><img src="/admin/images/caixa/top_esq.gif" width="14" height="12" /></td>
          <td background="/admin/images/caixa/top_centro.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td width="14" height="12"><img src="/admin/images/caixa/top_dir.gif" width="14" height="12" /></td>
        </tr>
        <tr>
          <td background="/admin/images/caixa/esq.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td>
          
          <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
            <form id="buscar" name="buscar" method="get" action="relatorio_ingressos.asp">
              <tr>
                <td height="22" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:14px; font-weight:bold;">Filtro</span></td>
              </tr>
              <tr>
                <td height="25" class="titulo_menu_site_bts" style="width:160px;"> Status do Pedidos:</td>
                <td class="titulo_noticias_home">
                    <select id="id_status" name="id_status" class="admin_txtfield_login">
                        <option value="">-- Todos os Pedidos --</option>
                        <%
                          If not RS_Status.BOF or not RS_Status.EOF Then
                            While not RS_Status.EOF
                                selecionado = ""
                                If Cstr(id_status) = Cstr(RS_Status("ID_Pedido_Status")) Then selecionado = " selected "
                                %><option value="<%=RS_Status("ID_Pedido_Status")%>" <%=selecionado%>><%=RS_Status("Status_PTB")%></option><%
                              RS_Status.MoveNext
                            Wend
                            RS_Status.Close
                          End If
                          %>
                    </select>
                </td>
              </tr>
              <tr>
              	<td class="titulo_menu_site_bts" style="height:25px;">Ano do Pedido:</td>
                <td>
                    <select name="ano_pedido" id="ano_pedido" class="admin_txtfield_login" style="width:145px">
                        <option value="">-- Todos os Anos --</option>
                        <%  '//Lista os anos, do ano atual até 2010
                            For ano = Year(Now()) To 2010 Step -1
                                If CStr(ano) = CStr(ano_pedido) Then
                                    Response.Write("<option selected>" & ano & "</option>")
                                Else
                                    Response.Write("<option>" & ano & "</option>")
                                End If
                            Next
                        %>
                    </select>
                </td>
              </tr>
              <tr>
                <td height="25" class="titulo_menu_site_bts"> Busca Personalizada:</td>
                <td class="titulo_noticias_home">
                	<input type="text" name="pedido" id="pedido" class="admin_txtfield_login" size="20" value="<%=pedido%>" style="width:145px"
						/> <select id="campo_busca" name="campo_busca" class="admin_txtfield_login" style="margin-top:5px; width:125px">
                        <option value="numeropedido">Número de Pedido</option>
                        <option value="nomecomprador">Nome do Comprador</option>
                        <option value="cpfcomprador">CPF do Comprador</option>
                    </select>
                </td>
              </tr>
			  <tr>
                <td height="25" class="titulo_menu_site_bts"> Ordenar Por:</td>
                <td class="titulo_noticias_home">
                    <select name="campo_ordem" id="campo_ordem" class="admin_txtfield_login" style="width:145px">
                        <option value="">-- Selecione --</option>
						<option value="datapedido">Data do Pedido</option>
						<option value="nomecomprador">Nome do Comprador</option>
						<option value="numeropedido">Número do Pedido</option>
                    </select>
					<select name="tipo_ordem" id="tipo_ordem" class="admin_txtfield_login" style="width:125px">
						<option value="">-- Selecione --</option>
						<option value="crescente">Crescente</option>
						<option value="decrescente">Decrescente</option>
					</select>
                </td>
			  </tr>
              <tr>
                <td height="30" class="titulo_menu_site_tec">&nbsp;</td>
                <td class="titulo_noticias_home">
                  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:left;"></div>
                  <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco cursor" onClick="document.buscar.submit()">Gerar Relatório</div>
                  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:right; width:15px; height:17px; float:left; margin-right:10px;"></div>
				  
				  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:left;"></div>
                  <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco cursor" onClick="Exportar_Excel()">Exportar Relatório</div>
                  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:right; width:15px; height:17px; float:left;"></div>
                  </td>
                </tr>
				<input type="hidden" name="ID_Evento" value="<%=ID_Evento%>">
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
      
      <br>
      
      <table width="900" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="14" height="12"><img src="/admin/images/caixa/top_esq.gif" width="14" height="12" /></td>
          <td background="/admin/images/caixa/top_centro.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td width="14" height="12"><img src="/admin/images/caixa/top_dir.gif" width="14" height="12" /></td>
        </tr>
        <tr>
          <td background="/admin/images/caixa/esq.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td>
          
          	<table>
            	<tr>
                	<td><img src="/admin/images/icones/Novo_Pedido.png"></td>
                    <td width="150" class="titulo_noticias_home">Novo Pedido</td>
                    
                    <td><img src="/admin/images/icones/Pedido_Concluido.png"></td>
                    <td width="200" class="titulo_noticias_home">Pedido Concluído</td>
                    
                    <td><img src="/admin/images/icones/Pedido_Cancelado.png"></td>
                    <td class="titulo_noticias_home">Pedido Cancelado</td>
                </tr>
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
      
      <br>
      
      <table width="900" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="14" height="12"><img src="/admin/images/caixa/top_esq.gif" width="14" height="12" /></td>
          <td background="/admin/images/caixa/top_centro.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td width="14" height="12"><img src="/admin/images/caixa/top_dir.gif" width="14" height="12" /></td>
        </tr>
        <tr>
          <td background="/admin/images/caixa/esq.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td>
          
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td height="30" colspan="2" align="center">
                
                <!--#include virtual="/admin/inc/paginacao.asp"-->
                
                <%
				Contador = 0
				%>
                
                	<table style="font-family:Arial, Helvetica, sans-serif; font-size:12px;" cellpadding="3" cellspacing="3" width="100%">
                    	<tr>
                        	<td class="borda_dir linha_16px" width="10"></td>
                            <td class="borda_dir linha_16px" width="130"><strong>Número do Pedido</strong></td>
                            <td class="borda_dir linha_16px" width="*"><strong>Nome do Comprador</strong></td>
                            <td class="borda_dir linha_16px" align="center" width="90"><strong>CPF</strong></td>
                            <td class="borda_dir linha_16px" align="center" width="70"><strong>Ingressos</strong></td>
                            <td class="borda_dir linha_16px" align="center" width="70"><strong>Valor</strong></td>
                            <td class="borda_dir linha_16px" align="center" width="70"><strong>Data</strong></td>
                            <td class="borda_dir linha_16px" align="center" width="40"><strong>Status</strong></td>
                            <td class="borda_dir linha_16px" align="center" width="40"><strong>Ação</strong></td>
                        </tr>
                        <%
						P = 0
						
						If Not RS_Listar.EOF Then
						
							RS_Listar.MoveFirst
							RS_Listar.AbsolutePage = PaginaAtual 
							
							While Not RS_Listar.EOF And Contador < RS_Listar.PageSize
							Contador = Contador + 1	
							
							P = P + 1
							%>
							<tr>
								<td class="borda_dir linha_16px"><%=P%></td>
								<td class="borda_dir linha_16px"><%=RS_Listar("Numero_Pedido")%></td>
								<td class="borda_dir linha_16px"><%=RS_Listar("Nome_Comprador")%></td>
								<td class="borda_dir linha_16px" align="center"><%=RS_Listar("CPF_Comprador")%></td>
								<td class="borda_dir linha_16px" align="center"><%=RS_Listar("Quantidade")%></td>
								<td class="borda_dir linha_16px" align="right">R$ <%=FormatNumber(RS_Listar("Valor_Pedido"))%></td>
								<td class="borda_dir linha_16px"><%
									If Day(RS_Listar("Data_Pedido")) < 10 Then
										Dia = "0" & Day(RS_Listar("Data_Pedido"))
									Else
										Dia = Day(RS_Listar("Data_Pedido"))
									End If
									
									If Month(RS_Listar("Data_Pedido")) < 10 Then
										Mes = "0" & Month(RS_Listar("Data_Pedido"))
									Else
										Mes = Month(RS_Listar("Data_Pedido"))
									End If
									
									
									
									Response.Write(Dia &"/"& Mes &"/"& Year(RS_Listar("Data_Pedido")))
									%></td>
								<td class="borda_dir linha_16px" align="center">
								<%
								
								If Cint(RS_Listar("ID_Status")) = 1 Then
								%><img src="/admin/images/icones/Novo_Pedido.png"><%
								ElseIf Cint(RS_Listar("ID_Status")) = 3 Then
								%><img src="/admin/images/icones/Pedido_Concluido.png"><%
								ElseIf Cint(RS_Listar("ID_Status")) = 4 Then
								%><img src="/admin/images/icones/Pedido_Cancelado.png"><%
								End If
								
								%>
								</td>
								<td class="borda_dir linha_16px" align="center">
									<a href="#detalhes" onClick="cancelar_pedido('<%=RS_Listar("Numero_Pedido")%>')"><img src="/admin/images/ico_desativado.gif" alt="Cancelar Pedido" title="Cancelar Pedido"></a><br>
									<a href="#detalhes" onClick="$('#<%=RS_Listar("Numero_Pedido")%>').toggle()"><img src="/admin/images/ico_busca_24px.gif" alt="Detalhes" title="Detalhes"></a>
								</td>
							</tr>
                            <tr id="<%=RS_Listar("Numero_Pedido")%>" style="display: none">
                            	<td colspan="9">
                                
                                <table style="font-family:Arial, Helvetica, sans-serif; font-size:12px;" cellpadding="3" cellspacing="3" width="100%" bgcolor="#E8E8E8">
                                <%
								SQL_Carrinho = 	"Select " &_
												"	Vi.Nome_Completo " &_
												"	,Vi.CPF " &_
												"	,Vi.Passaporte " &_
												"	,Vi.Email " &_
												"From Pedidos_Carrinho As Pc " &_
												"Inner Join Relacionamento_Cadastro As Rc " &_
												"	On Pc.ID_Rel_Cadastro = Rc.ID_Relacionamento_Cadastro " &_
												"		And Pc.ID_Visitante = Rc.ID_Visitante " &_
												"Inner Join Visitantes As Vi " &_
												"	On Pc.ID_Visitante = Vi.ID_Visitante " &_
												"		And Rc.ID_Visitante = Vi.ID_Visitante " &_
												"Where ID_Pedido = '" & RS_Listar("ID_Pedido") & "'"
												
												'Response.Write SQL_Carrinho
												'response.end
								Set RS_Carrinho = Server.CreateObject("ADODB.Recordset")
								RS_Carrinho.Open SQL_Carrinho, Conexao, 3, 3
								
								If Not RS_Carrinho.Eof Then
									
									While Not RS_Carrinho.Eof
									
									%>
                                	<tr>
                                    	<td class="borda_dir linha_16px" width="418"><%=RS_Carrinho("Nome_Completo")%></td>
                                        <td class="borda_dir linha_16px"><%If Len(Trim(RS_Carrinho("CPF"))) <> 0 Then Response.Write(RS_Carrinho("CPF")) Else Response.Write(RS_Carrinho("Passaporte"))%></td>
                                        <td class="borda_dir linha_16px"><%=RS_Carrinho("Email")%></td>
                                    </tr>
									<%
									
									RS_Carrinho.MoveNext
									Wend
									
								End If
								%>
                                </table>
                                </td>
                            </tr>
							<%
							RS_Listar.MoveNext
							Wend
						
						Else
						%>
                    	<tr>
                        	<td class="borda_dir linha_16px" colspan="9">
                            	Não existem resultados com esta informação!
                            </td>
                        </tr>
						<%
						End If
						%>
                    </table>
                    
                    <!--#include virtual="/admin/inc/paginacao.asp"-->
                
                </td>
              </tr>
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