<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
id	= Limpar_Texto(Request("id_edicao"))

'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")

	SQL_Eventos = 	"Select " &_
					"	Ee.ID_Edicao, " &_
					"	E.Nome_PTB as Evento, " &_
					"	Ee.Ano " &_
					"From Eventos_Edicoes as Ee " &_
					"Inner Join Eventos as E ON E.ID_Evento = Ee.ID_Evento " &_
					"Where " &_
					"	ID_Edicao = " & id
					
	Set RS_Eventos = Server.CreateObject("ADODB.Recordset")
	RS_Eventos.Open SQL_Eventos, Conexao
	
	If RS_Eventos.BOF or RS_Eventos.EOF Then
		response.Redirect("default.asp?erro=id_nao_encontrado")
	Else
		Evento = RS_Eventos("Ano") & " - " & RS_Eventos("Evento")
		RS_Eventos.Close
	End If

	SQL_SubCargos = "Select " &_
					"	S.ID_SubCargo " &_
					"	,S.SubCargo_PTB " &_
					"	,C.Cargo_PTB " &_
					"	,S.Ativo " &_
					"From SubCargo as S " &_
					"Inner Join Cargo as C ON C.ID_Cargo = S.ID_Cargo " &_
					"Order by Cargo_PTB, SubCargo_PTB"

	Set RS_SubCargos = Server.CreateObject("ADODB.Recordset")
	RS_SubCargos.Open SQL_SubCargos, Conexao

	SQL_Listar = 	"Select " &_
					"	ID_SubCargo " &_
					"From Relacionamento_Edicoes_SubCargo as Rec " &_
					"Where " &_
					"	ID_Edicao = " & id

'	response.write("<hr>" & SQL_Listar & "<hr>")

	Set RS_Listar = Server.CreateObject("ADODB.Recordset")
	RS_Listar.Open SQL_Listar, Conexao
	
	If RS_Listar.BOF or RS_Listar.EOF Then
		'response.Redirect("default.asp?msg=erro_nao_encontrado")
		registros = 0
		Redim IDs(registros)
	Else
		registros = 0
		While not RS_Listar.EOF
			registros = registros + 1
			RS_Listar.MoveNext
		Wend
		RS_Listar.MoveFirst
		Redim IDs(registros)
		
		x = 0
		While not RS_Listar.EOF
			IDs(x) = RS_Listar("ID_SubCargo")
			x = x + 1
			RS_Listar.MoveNext
		Wend
		RS_Listar.Close
	End If
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
			$('#aviso_conteudo').html('Evento adicionado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
			<%
		Case "add_erro_existe"
			%>
			$('#aviso_conteudo').html('Erro - Evento informado já existe !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
			<%
		Case "upd_ok"
			%>
			$('#aviso_conteudo').html('Evento atualizado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
			<%
		Case "atv_ok"
			%>
			$('#aviso_conteudo').html('Evento ativado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
			<%
		Case "des_ok"
			%>
			$('#aviso_conteudo').html('Evento desativado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
			<%
	End Select
	%>
});

function voltar() {
	confirmacao = confirm("Os dados não foram salvos, deseja sair ?")
	if (confirmacao) {
		document.location = 'default.asp';	
	}
}

function check (qual) {
	if ( $(':checkbox')[ + qual ].checked ) {
		$(':checkbox')[ + qual ].checked = false;
	} else {
		$(':checkbox')[ + qual ].checked = true;
	}
}


function Enviar() {
	var n = $("input:checked").length;
	var enviando = false;
	if (n == 0) {
		var erros = 1;
	} else {
		var erros = 0;
	}
	//alert( erros );
	if (erros == 0 && enviando == false) {
		enviando = true;
		
		// Alterar Imagem Src / W / H
		$('#ico_salvar').attr('src','/admin/images/ico_loading1.gif');
		$('#ico_salvar').attr('width','16');
		$('#ico_salvar').attr('height','16');
		
		// Alterar texto + criar div para processamento
		$('#txt_salvar').fadeOut().html('Enviando dados...').fadeIn();
		
		// Realizar POST 'comum'
		document.cad.submit();
		
	} else if (erros == 0 && enviando == true) {
		//alert('Processando, por favor aguarde !');
		$('#txt_salvar').fadeOut().fadeIn();
	} else {
		alert('Selecione ao Menos 1 ITEM !');
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
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">Relacionar SubCargos - Edi&ccedil;&otilde;es<br>
          Edi&ccedil;&atilde;o: <font color="#0000FF"><%=Evento%></font></span></td>
        <td width="100" align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;"><a href="javascript:voltar();"><img src="/admin/images/bt_voltar_admin.gif" width="100" height="48" border="0" /></a><a href="default.asp"></a></td>
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
          <td>
          <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td width="260" height="30" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">Marque os SubCargos Desejados</span></td>
                </tr>
              <tr>
                <td colspan="2">
                
                    <form name="cad" id="cad" method="post" action="metodos.asp">
                      <div id="forms_hidden" style="position:fixed; left:-100px; top:-100px;">
                        <input type="hidden" id="acao" name="acao" value="menus_relacionar" style="visibility:hidden; display:none;">
                        <input type="hidden" id="id_edicao" name="id_edicao" value="<%=id%>" style="visibility:hidden; display:none;">
                      </div>
                      <table width="420" border="0" cellpadding="0" cellspacing="5">
                        <tr>
                          <td width="180" height="230" valign="top" bgcolor="#E3E8FD" class="admin_tela_login">
                          <table width="390" border="0" cellspacing="0" cellpadding="0" class="fs11px t_arial bold c_vermelho">
                            <tr>
                              <td colspan="2" align="center">SubCargos</td>
                            </tr>
                          </table>
                            <div style="height:380px; overflow:auto;" id="lista_contato">
                              <table id="lista_menus" width="390" border="0" cellspacing="0" cellpadding="0" class="admin_tela_login">
                                <%
								If not RS_SubCargos.BOF or RS_SubCargos.EOF Then
									i = 0
									While not RS_SubCargos.EOF
										checado = ""
										For w = Lbound(IDs) to Ubound(IDs)
											If Cstr(IDs(w)) = Cstr(RS_SubCargos("id_SubCargo")) Then 
												checado = " checked "
												Exit For
											End If
										Next
										ativo = RS_SubCargos("ativo")
										txt_desativado = ""
										If ativo = false OR ativo = 0 Then txt_desativado = "<span style='color: #f00'> (desativado) </span>"
										%>
                                        <tr onMouseOver="$(this).addClass('bg_radio');" onMouseOut="$(this).removeClass('bg_radio');">
										<td width="30" height="24" align="center"><input type="checkbox" name="id_SubCargo" id="id_SubCargo" value="<%=RS_SubCargos("id_SubCargo")%>" <%=checado%>></td>
                                        <td onClick="check(<%=i%>);" class="cursor"><span style="color: #669;"><%=RS_SubCargos("Cargo_PTB")%></span> - <%=RS_SubCargos("SubCargo_PTB")%> <%=txt_desativado%></td>
                                        </tr>
                                    	<%
										RS_SubCargos.MoveNext
										i = i + 1
									Wend
									RS_SubCargos.Close
								End If
								%>
                              </table>
                            </div>
                            </td>
                        </tr>
                      </table>
                      <table width="420" border="0" cellpadding="0" cellspacing="5">
                        <tr id="salvar">
                          <td width="130" height="25" align="right" class="admin_tela_login"><img src="/admin/images/ico_salvar.gif" alt="Gravar Alterações" name="ico_salvar" width="20" height="20" id="ico_salvar" title="Gravar Alterações"></td>
                          <td class="titulo_menu_site_carne"><span style="cursor:pointer;" onClick="Enviar();" id="txt_salvar">Salvar Menus</span></td>
                          <td width="25" height="25" align="center"></td>
                          <td width="25" align="center">&nbsp;</td>
                        </tr>
                      </table>
                    </form>
                
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
      <table width="900" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td>&nbsp;</td>
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