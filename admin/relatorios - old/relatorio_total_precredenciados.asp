<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
' * Dados Paginação
qtde					= Limpar_Texto(Request("qtd"))
ordem					= Limpar_Texto(Request("ord"))
acao_pagina		= Limpar_Texto(Request("acp"))
pagina				= Limpar_Texto(Request("pag"))

strURL = "relatorio_precredenciados.asp?w=0"

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

' * Request BUSCA
id_edicao         = Limpar_Texto(Request("id_edicao"))

If id_edicao <> "" and id_edicao <> "-" Then

  SQL_Listar =  "Select " &_
                " Count(TC.ID_Formulario) as Total " &_
                "    ,F.ID_Formulario as ID_Formulario " &_
                "    ,F.Nome as Formulario " &_
                "   ,I.Nome as Idioma " &_
                "   ,RC.ID_Idioma as ID_Idioma " &_
                "  From  " &_
                "    Relacionamento_Cadastro as RC  " &_
                "  Inner Join  " &_
                "    Tipo_Credenciamento as TC  " &_
                "    ON TC.ID_Tipo_Credenciamento = RC.ID_Tipo_Credenciamento  " &_
                "  Inner Join  " &_
                "    Formularios as F  " &_
                "    ON F.ID_Formulario = TC.ID_Formulario  " &_
                "  Inner Join  " &_
                "    Eventos_Edicoes as EE  " &_
                "    ON EE.ID_Edicao = RC.ID_Edicao " &_
                "  Inner Join  " &_
                "    Eventos as EV  " &_
                "    ON EV.ID_Evento = EE.ID_Evento  " &_
                "  Inner Join  " &_
                "    Visitantes as V  " &_
                "    ON V.ID_Visitante = RC.ID_Visitante  " &_
                "  Inner Join " &_
                "    Idiomas as I " &_
                "    ON I.ID_Idioma = RC.ID_Idioma " &_
                "  Left Join  " &_
                "    Empresas as E  " &_
                "    ON E.ID_Empresa = RC.ID_Empresa  " &_
                " Where RC.ID_Edicao = " & id_edicao & " " &_
                "   Group By F.Nome, F.ID_Formulario, I.Nome, RC.ID_Idioma " &_
                "   Order by Idioma "
  'response.write("<hr>" & SQL_Listar & "<hr>")

  Set RS_Listar = Server.CreateObject("ADODB.Recordset")
  'RS_Listar.CursorLocation = 3
  RS_Listar.Open SQL_Listar, Conexao

End If
	
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
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">Listagem de Pr&eacute;-Credenciados</span></td>
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
            <form id="buscar" name="buscar" method="post">
              <tr>
                <td height="30" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">Buscar Pr&eacute;-Credenciado</span></td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_bts"> Edi&ccedil;&atilde;o</td>
                <td class="titulo_noticias_home">
                <select id="id_edicao" name="id_edicao" class="admin_txtfield_login" onChange="document.location='?id_edicao='+this.value;">
                <option value="-">-- Selecione --</option>
                <%
                  If not RS_Eventos.BOF or not RS_Eventos.EOF Then
                    While not RS_Eventos.EOF
                      selecionado = ""
                      If Cstr(id_edicao) = Cstr(RS_Eventos("ID_Edicao")) Then
                        selecionado = " selected "
                      End If
                    %><option <%=selecionado%> value="<%=RS_Eventos("ID_Edicao")%>"><%=RS_Eventos("ano")%> - <%=RS_Eventos("Evento")%></option><%
                      RS_Eventos.MoveNext
                    Wend
                    RS_Eventos.Close
                  End If
                  %>
                </select>
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
      <table width="450" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="14" height="12"><img src="/admin/images/caixa/top_esq.gif" width="14" height="12"></td>
          <td background="/admin/images/caixa/top_centro.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12"></td>
          <td width="14" height="12"><img src="/admin/images/caixa/top_dir.gif" width="14" height="12"></td>
        </tr>
        <tr>
          <td background="/admin/images/caixa/esq.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12"></td>
          <td>
            <%	
              If id_edicao <> "" and id_edicao <> "-" Then
                If Rs_listar.BOF or Rs_listar.EOF Then	%>
                  <p align="center" class="titulo_menu_site_carne">N&atilde;o foi encontrado nenhum registro</p>
            <% End If %>
            <%
              If not Rs_listar.BOF or not Rs_listar.EOF Then
        		%>
            <table width="100%" border="0" cellpadding="2" cellspacing="1" class="fs11px t_arial">
              <tr><td id="titulo_feira" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold; color: #00F">
                <script language="javascript">
                  texto = $("select[name=id_edicao] option:selected").text();
                  document.write(texto);
                </script>
              </span>
              </td></tr>
            </table>    
            <table width="100%" border="0" cellpadding="2" cellspacing="1" class="fs11px t_arial">
              <tr>
                <td class="borda_dir linha_16px" align="center"><b>Formularios</b></td>
                <td class="borda_dir linha_16px" align="center"><b>Quantidade</b></td>
              </tr>
              <%
			Total_Geral = 0
            Rs_listar.MoveFirst
            While Not RS_Listar.EOF
              if id_edicao = 1 Then ' CARDS 2012 
                MsgTotal = "Existem dados no Banco Antigo e no Banco Novo"
                Select Case RS_listar("ID_Formulario")
                  Case "1"
                    if RS_listar("ID_Idioma") = "1" Then
                      Total = RS_Listar("Total") + 284
                    Else
                      Total = RS_Listar("Total")
                    End If
                  Case "2"
                    if RS_listar("ID_Idioma") = "1" Then
                      Total = RS_Listar("Total") + 4
                    Else
                      Total = RS_Listar("Total")
                    End If
                  Case "3"  
                    Total = RS_Listar("Total")
                  Case "4"  
                    Total = RS_Listar("Total") 
                  Case "5"  
                    Total = RS_Listar("Total")
                End Select
              Elseif id_edicao = 2 Then ' FORMOBILE 2012 
                MsgTotal = "Existem dados no Banco Antigo e no Banco Novo"
                Select Case RS_listar("ID_Formulario")
                  Case "1"
                    if RS_listar("ID_Idioma") = "1" Then
                      Total = RS_Listar("Total") + 1876
                    Else
                      Total = RS_Listar("Total")
                    End If
                  Case "2"
                    Total = RS_Listar("Total")
                  Case "3"  
                    Total = RS_Listar("Total")
                  Case "4"  
                    Total = RS_Listar("Total") 
                  Case "5"  
                    Total = RS_Listar("Total")
                End Select
              Else
                Total = ""
                Total = RS_Listar("Total")
              End If
			  Total_Geral = Total_Geral + Total
            %>
              <tr>
                <td class="borda_dir linha_16px" align="center"><b><%=RS_listar("Formulario")%> - <%=RS_listar("Idioma")%></b></td>
                <td class="borda_dir linha_16px" align="center"><%=Total%></td>
              </tr>
            <%
                RS_Listar.MoveNext()
            Wend
            RS_Listar.Close
			%>
            </table>
            <br/><span class="fs11px t_arial" align="center" style="color:#f00">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b><%=MsgTotal%></b></span>
            <%
          End If
        End If
        %>
        	<br><br><b><span class="fs11px t_arial">Total_Geral: <%=Total_Geral%></span></b>
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