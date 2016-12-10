<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
' * Dados Paginação
qtde					= Limpar_Texto(Request("qtd"))
ordem					= Limpar_Texto(Request("ord"))
acao_pagina				= Limpar_Texto(Request("acp"))
pagina					= Limpar_Texto(Request("pag"))

id	= Limpar_Texto(Request("id"))

strURL = "default.asp?id=" & id

If Len(qtde) > 0 AND isNumeric(qtde) Then 
	qtde_itens = qtde
	strURL = strURL & "&qtd=" & qtde
Else
	qtde_itens = 50
	qtde = 50
End If

'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")

	SQL_Eventos = "Select " &_
					"	* " &_
					"From Eventos_Edicoes AS Ee " &_
					"	Inner Join Eventos AS E" &_
					"	ON Ee.ID_Evento = E.ID_Evento Order by E.ID_EVENTO DESC"
	Set RS_Eventos = Server.CreateObject("ADODB.Recordset")
	RS_Eventos.Open SQL_Eventos, Conexao

	SQL_Tipo = 		"Select " &_
					"	ID_Tipo, " &_
					"	Tipo " &_
					"From Paginas_Tipo "
	Set RS_Tipo = Server.CreateObject("ADODB.Recordset")
	RS_Tipo.Open SQL_Tipo, Conexao
	
	SQL_PaginaWeb = "Select " &_
					"	Pagina " &_
					"From Paginas_Web " &_
					"Where ID_Pagina = " & id
	Set RS_PaginaWeb = Server.CreateObject("ADODB.Recordset")
	RS_PaginaWeb.Open SQL_PaginaWeb, Conexao
	
	pagina_nome = RS_PaginaWeb("pagina")
	RS_PaginaWeb.Close
	
	SQL_Ordem =		"Select " &_
					"	Max(Ordem) + 1 as Ordem " &_
					"From Paginas_Textos " &_
					"Where ID_Pagina = " & id
	Set RS_Ordem = Server.CreateObject("ADODB.Recordset")
	RS_Ordem.Open SQL_Ordem, Conexao
	
	nova_ordem = 1
	If not RS_Ordem.BOF OR RS_Ordem.EOF Then
		nova_ordem = RS_Ordem("ordem")
		If isNull(nova_ordem) Then nova_ordem = 1
		RS_Ordem.Close
	End If

'	SQL_Listar = 	"Select " &_
'					"	Pt.ID_Texto, " &_
'					"	I.Nome as Idioma, " &_
'					"	A.Nome as Administrador, " &_
'					"	T.Tipo, " &_
'					"	Pt.Ordem, " &_
'					"	Pt.Identificacao, " &_
'					"	Pt.Texto, " &_
'					"	Pt.URL_Imagem, " &_
'					"	Pt.Data_Cadastro as Data, " &_
'					"	Pt.Data_Atualizacao " &_
'					"From Paginas_Textos as Pt " &_
'					"Inner Join Idiomas as I ON I.ID_Idioma = Pt.ID_Idioma " &_
'					"Inner Join Paginas_Tipo as T ON T.ID_Tipo = Pt.ID_Tipo " &_
'					"Inner Join Administradores as A ON A.ID_Admin = Pt.ID_Admin " &_
'					"Where " &_
'					"	Pt.ID_Pagina = " & id & " " &_

	SQL_Listar =	"Select Distinct " &_
					"	A.Nome as Administrador, " &_
					"	Pt.Ordem, " &_
					"	Pt.Identificacao " &_
					"From Paginas_Textos as PT " &_
					"Inner Join Administradores as A ON A.ID_Admin = Pt.ID_Admin " &_
					"Where	ID_Pagina = " & id & " " &_
					"Order by Ordem DESC"

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
<title>Administração CSC - BTS Informa</title>
<link href="/admin/css/bts.css" rel="stylesheet" type="text/css" />
<link href="/admin/css/admin.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
<script language="javascript" src="/admin/js/validar_forms.js"></script>
</head>


<script language="javascript">
$(document).ready(function(){
	$('#aviso').hide();
	$('#protheus').select().focus();
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
			case 'texto_ptb':
				break;
			case 'texto_eng':
				break;
			case 'texto_esp':
				break;
			case 'img_ptb':
				break;
			case 'img_eng':
				break;
			case 'img_esp':
				break;
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
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center">P&aacute;gina: <%=pagina_nome%><br>
          <span style="color: #B01D22">Textos das P&aacute;ginas</span></td>
        <td width="100" align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;"><a href="/admin/paginas_web/default.asp"><img src="/admin/images/bt_voltar_admin.gif" width="100" height="48" border="0" /></a></td>
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
            <input type="hidden" id="acao" name="acao" value="add_texto">
            <input type="hidden" id="id_pagina" name="id_pagina" value="<%=id%>">
              <tr>
                <td height="30" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">Cadastrar Novo Texto</span></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts"> Evento</td>
                <td class="titulo_noticias_home"><select id="id_evento" name="id_evento" class="admin_txtfield_login">
                  <option value="-">-- Selecione --</option>
                  <option value="Null">Geral</option>
                  <%
							If not RS_Eventos.BOF or not RS_Eventos.EOF Then
								While not RS_Eventos.EOF
						%>
                  <option value="<%=RS_Eventos("ID_Edicao")%>"><%=RS_Eventos("ID_Edicao")%> - <%=RS_Eventos("Ano")%> - <%=RS_Eventos("Nome_PTB")%></option>
                  <%
									RS_Eventos.MoveNext()
								Wend
								RS_Eventos.Close  
							End If						
						%>
                </select></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts"> Tipo de Tradu&ccedil;&atilde;o:</td>
                <td class="titulo_noticias_home"><select id="id_tipo" name="id_tipo" class="admin_txtfield_login">
                  <option value="-">-- Selecione --</option>
                  <%
				If not RS_Tipo.BOF or not RS_Tipo.EOF Then
					While not RS_Tipo.EOF
					%>
                  <option value="<%=RS_Tipo("ID_Tipo")%>"><%=RS_Tipo("Tipo")%></option>
                  <%
						RS_Tipo.MoveNext
					Wend
					RS_Tipo.Close
				End If
				%>
                  </select></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Ordem:</td>
                <td class="titulo_noticias_home"><%=nova_ordem%><input type="hidden" id="ordem" name="ordem" value="<%=nova_ordem%>"></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Identifica&ccedil;&atilde;o:</td>
                <td class="titulo_noticias_home"><input name="identificacao" id="identificacao" type="text" class="admin_txtfield_login" size="30" /></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Texto PTB</td>
                <td class="titulo_noticias_home"><textarea name="texto_ptb" cols="30" rows="3" class="admin_txtfield_login" id="texto_ptb"></textarea></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Texto ENG</td>
                <td class="titulo_noticias_home"><textarea name="texto_eng" cols="30" rows="3" class="admin_txtfield_login" id="texto_eng"></textarea></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Texto ESP</td>
                <td class="titulo_noticias_home"><textarea name="texto_esp" cols="30" rows="3" class="admin_txtfield_login" id="texto_esp"></textarea></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Url Imagem PTB</td>
                <td class="titulo_noticias_home"><input name="img_ptb" id="img_ptb" type="text" class="admin_txtfield_login" size="30"  value="/img/" />
                  <img src="/admin/images/ico_preview_20_b.gif" alt="Visualizar Imagem" title="Visualizar Imagem" class="cursor" onClick="window.open($('#logo_faixa').val());" width="20" height="20" align="absmiddle"></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Url Imagem ENG</td>
                <td class="titulo_noticias_home"><input name="img_eng" id="imag_eng" type="text" class="admin_txtfield_login" size="30"  value="/img/" />
                  <img src="/admin/images/ico_preview_20_b.gif" alt="Visualizar Imagem" title="Visualizar Imagem" class="cursor" onClick="window.open($('#logo_faixa').val());" width="20" height="20" align="absmiddle"></td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Url Imagem ESP</td>
                <td class="titulo_noticias_home"><input name="img_esp" id="img_esp" type="text" class="admin_txtfield_login" size="30"  value="/img/" />
                  <img src="/admin/images/ico_preview_20_b.gif" alt="Visualizar Imagem" title="Visualizar Imagem" class="cursor" onClick="window.open($('#logo_faixa').val());" width="20" height="20" align="absmiddle"></td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_tec">&nbsp;</td>
                <td class="titulo_noticias_home">
                  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:left;"></div>
                  <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco cursor" onClick="Enviar();">Cadastrar Novo Texto</div>
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
                <td width="50" align="center" class="borda_dir linha_16px"><b>Ordem</b></td>
                <td align="center" class="borda_dir linha_16px"><b>Identificacao</b></td>
                <td width="150" align="center" class="borda_dir linha_16px"><b>Atualizado por - Data</b></td>
                <td width="40" align="center" class="linha_16px"><b>Editar</b></td>
              </tr>
              <tr>
              	<td colspan="5" bgcolor="#000000" height="3"><img src="/img/geral/spacer.gif" width="1" height="3"></td>
              </tr>
            <%
            Rs_listar.MoveFirst
            RS_Listar.AbsolutePage = PaginaAtual 
            While Not RS_Listar.EOF And Contador < RS_Listar.PageSize
            Contador = Contador + 1
			
				SQL_Data = 	"Select " &_
							"	Data_Cadastro, " &_
							"	Data_Atualizacao " &_
							"From Paginas_Textos " &_
							"Where Ordem = " & Rs_listar("ordem")
				Set RS_Data = Server.CreateObject("ADODB.Recordset")
				RS_Data.Open SQL_Data, Conexao
				
				Data_Cadastro = RS_Data("Data_Cadastro")
				Data_Atualizacao = RS_Data("Data_Atualizacao")
				RS_Data.Close
            %>
              <tr bgcolor="#FFFFFF" 
                onMouseOver="$(this).attr('bgcolor','#FFFF00'); " 
                onMouseOut="$(this).attr('bgcolor','#FFFFFF'); "  >
                <td class="borda_dir linha_16px fs12px" align="center"><strong><%=Rs_listar("Ordem")%></strong></td>
                <td class="borda_dir linha_16px" align="center" bgcolor="#CCCCCC"><%=Rs_listar("Identificacao")%></td>
                <td class="borda_dir linha_16px" align="left">
                <%=RS_Listar("Administrador")%> - 
				<%
				If Data_Atualizacao <> "" Then
					%><%=FormatDateTime(Data_Atualizacao,2)%><%
				Else
					%><%=FormatDateTime(Data_Cadastro,2)%><%
				End If
				%>
                </td>
                <td class="borda_dir linha_16px" align="center"><a href="editar.asp?idp=<%=id%>&ord=<%=Rs_listar("ordem")%>"><img src="/admin/images/ico_editar.gif" width="20" height="20" alt="Editar Página" title="Editar Página" border="0"></a></td>
              </tr>
              <tr>
              	<td colspan="5">
				<table width="99%" cellpadding="2" cellspacing="2" class="fs11px t_arial" bgcolor="#F1F1F1" align="center">
                    <tr>
                        <td align="center" class="borda_dir linha_16px" width="100"><strong>Idioma</strong></td>
                        <td align="center" class="borda_dir linha_16px"><strong>Texto</strong></td>
                        <td align="center" class="borda_dir linha_16px" width="33%"><strong>Imagem</strong></td>
                    </tr>
                    <%
					SQL_Textos = 	"Select " &_
									"	I.Nome as Idioma, " &_
									"	Pt.Texto, " &_
									"	PT.URL_Imagem " &_
									"From Paginas_Textos as Pt " &_
									"Inner Join Idiomas as I ON I.ID_Idioma = Pt.ID_Idioma " &_
									"Where " &_
									"	ID_Pagina = " & id & " " &_
									"	AND Ordem = " & Rs_listar("ordem")
					Set RS_Textos = Server.CreateObject("ADODB.Recordset")
					RS_Textos.Open SQL_Textos, Conexao,1
					
					If not RS_Textos.BOF or not RS_Textos.EOF Then
						While not RS_Textos.EOF
					%>
                    <tr>
                        <td align="left" class="borda_dir linha_16px" width="100"><%=RS_Textos("Idioma")%></td>
                        <td align="left" class="borda_dir linha_16px"><%=RS_Textos("Texto")%>&nbsp;</td>
                        <td align="center" class="borda_dir linha_16px" width="33%">
                        <% If Len(Trim(RS_Textos("URL_Imagem"))) > 0 Then %>
                        <img src="<%=RS_Textos("URL_Imagem")%>">
                        <% End If %>
                        &nbsp;</td>
                    </tr>
                    <%
							RS_Textos.MoveNext
						Wend
					End If
					%>
                </table>
                </td> 
              </tr>
              <tr>
              	<td colspan="5" bgcolor="#000000" height="3"><img src="/img/geral/spacer.gif" width="1" height="3"></td>
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