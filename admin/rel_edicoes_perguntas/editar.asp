<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
id				= Limpar_Texto(Request("id_edicao"))
tp_formulario	= Limpar_Texto(Request("id_formulario"))

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

	SQL_Formularios = 	"Select " &_
					"	ID_Formulario " &_
					"	,Nome " &_
					"From Formularios "
					
	Set RS_Formularios = Server.CreateObject("ADODB.Recordset")
	RS_Formularios.Open SQL_Formularios, Conexao
	
	If RS_Formularios.BOF or RS_Formularios.EOF Then
		response.Redirect("default.asp?erro=id_nao_encontrado")
	End If

If tp_formulario <> "" then

	SQL_Perguntas = 	"Select " &_
					"	ID_Perguntas " &_
					"	,ID_Edicao " &_
					"	,ID_Formulario " &_
					"	,Pergunta_PTB " &_
					"	,Ativo " &_
					"From Perguntas " &_
					"Where " &_
					"	ID_Formulario = " & tp_formulario & " "  &_
					"Order by Pergunta_PTB"

	'response.write("<hr>SQL_Perguntas<hr>" & SQL_Perguntas & "<hr>")

	Set RS_Perguntas = Server.CreateObject("ADODB.Recordset")
	RS_Perguntas.Open SQL_Perguntas, Conexao

	SQL_Listar = 	"Select " &_
					"	ID_Pergunta " &_
					"From Relacionamento_Edicoes_Pergunta as Rep " &_
					"Where " &_
					"	ID_Edicao = " & id

	'response.write("<hr>SQL_Listar<hr>" & SQL_Listar & "<hr>")

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
			IDs(x) = RS_Listar("ID_Pergunta")
			x = x + 1
			RS_Listar.MoveNext
		Wend
		RS_Listar.Close

	End If

End If
%>
<html>
<head>
<meta http-equiv="Content-type" content="text/html; charset=iso-8859-1" />
<title>Administra��o CSC - BTS Informa</title>
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
			$('#aviso_conteudo').html('P�gina n�o permitida !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000);
			<%
		Case "add_ok"
			%>
			$('#aviso_conteudo').html('Evento adicionado !');
			$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000).fadeOut('slow');
			<%
		Case "add_erro_existe"
			%>
			$('#aviso_conteudo').html('Erro - Evento informado j� existe !');
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
	confirmacao = confirm("Os dados n�o foram salvos, deseja sair ?")
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
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">Relacionar Perguntas - Edi&ccedil;&otilde;es<br>
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
                <td height="30" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">Escolha o Formul�rio</span></td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_bts">Formul�rio:</td>
                <td class="titulo_noticias_home">
                <select id="id_formulario" name="id_formulario" class="admin_txtfield_login" onChange="javascript:document.location='?id_edicao=' + <%=id%> + '&id_formulario=' + this.value;">
                	<option value="-">-- Selecione --</option>
                <%
				If not RS_Formularios.BOF or not RS_Formularios.EOF Then
					While not RS_Formularios.EOF
					
					If Cstr(tp_formulario) = Cstr(RS_Formularios("ID_Formulario")) Then
						checado = "selected"
					End If
					%><option value="<%=RS_Formularios("ID_Formulario")%>" <%=checado%> ><%=RS_Formularios("Nome")%></option><%
						RS_Formularios.MoveNext
						checado = ""
					Wend
					RS_Formularios.Close
				End If
				%>
                </select>
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
      <br/>
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
                <td width="260" height="30" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">Formul�rio de Empresa > Marque as Perguntas Desejadas</span></td>
                </tr>
              <tr>
                <td colspan="2">
                
                    <form name="cad" id="cad" method="post" action="metodos.asp">
                      <div id="forms_hidden" style="position:fixed; left:-100px; top:-100px;">
                        <input type="hidden" id="acao" name="acao" value="menus_relacionar" style="visibility:hidden; display:none;">
                        <input type="hidden" id="id_edicao" name="id_edicao" value="<%=id%>" style="visibility:hidden; display:none;">
                        <input type="hidden" id="id_form" name="id_form" value="<%=tp_formulario%>" style="visibility:hidden; display:none;">
                      </div>
                      <table width="420" border="0" cellpadding="0" cellspacing="5">
                        <tr>
                          <td width="180" height="230" valign="top" bgcolor="#E3E8FD" class="admin_tela_login">
                          <table width="500" border="0" cellspacing="0" cellpadding="0" class="fs11px t_arial bold c_vermelho">
                            <tr>
                              <td colspan="2" align="center">Perguntas</td>
                            </tr>
                          </table>
                            <div style="height:500px; overflow:auto;" id="lista_contato">
                              <table id="lista_menus" width="500" border="0" cellspacing="0" cellpadding="0" class="admin_tela_login">
                                <%
					          	if tp_formulario <> "" Then
								If not RS_Perguntas.BOF or RS_Perguntas.EOF Then
									i = 0
									While not RS_Perguntas.EOF
										checado = ""

										For w = Lbound(IDs) to Ubound(IDs)

											If Cstr(IDs(w)) = Cstr(RS_Perguntas("ID_Perguntas")) Then 
												checado = " checked "
												Exit For
											End If
										Next
										ativo = RS_Perguntas("ativo")
										txt_desativado = ""
										If ativo = false OR ativo = 0 Then txt_desativado = "<span style='color: #f00'> (desativado) </span>"
										%>
                                        <tr onMouseOver="$(this).addClass('bg_radio');" onMouseOut="$(this).removeClass('bg_radio');">
										<td width="30" height="24" align="center"><input type="checkbox" name="id_perguntas" id="id_perguntas" value="<%=RS_Perguntas("ID_Perguntas")%>" <%=checado%>></td>
                                        <td onClick="check(<%=i%>);" class="cursor">(<%=RS_Perguntas("ID_Perguntas")%>) <%=RS_Perguntas("Pergunta_PTB")%> <%=txt_desativado%></td>
                                        </tr>
                                    	<%
										RS_Perguntas.MoveNext
										i = i + 1
									Wend
									RS_Perguntas.Close
								End If
								Else
								%>
								<div align="center"><br/><br/>Selecione o Tipo de Formul�rio </div>
								<%
								End If
								%>
                              </table>
                            </div>
                            </td>
                        </tr>
                      </table>
                      <table width="420" border="0" cellpadding="0" cellspacing="5">
                        <tr id="salvar">
                          <td width="130" height="25" align="right" class="admin_tela_login"><img src="/admin/images/ico_salvar.gif" alt="Gravar Altera��es" name="ico_salvar" width="20" height="20" id="ico_salvar" title="Gravar Altera��es"></td>
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