<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
id	= Limpar_Texto(Request("id"))

'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")

	SQL_Listar = 	"Select " &_
					"	* " &_
					"From Eventos_Edicoes " &_
					"Where ID_Edicao = " & id

'	response.write("<hr>" & SQL_Listar & "<hr>")

	Set RS_Listar = Server.CreateObject("ADODB.Recordset")
	RS_Listar.Open SQL_Listar, Conexao

	If RS_Listar.BOF or RS_Listar.EOF Then
		response.Redirect("default.asp?msg=erro_nao_encontrado")
	Else

		SQL_Eventos = 	"Select " &_
						"	ID_Evento, " &_
						"	Nome_PTB " &_
						"From Eventos " &_
						"Where ativo = 1 " &_
						"Order by Nome_PTB"

		Set RS_Eventos = Server.CreateObject("ADODB.Recordset")
		RS_Eventos.Open SQL_Eventos, Conexao

		id_evento 	= RS_Listar("id_evento")
		ano 		= RS_Listar("ano")
		inicio 		= RS_Listar("Data_Inicio_Feira")
		fim 		= RS_Listar("Data_fim_Feira")

		dia = Day(Now)
		If Len(dia) < 2 Then dia = "0" & dia
		mes = Month(Now)
		If Len(mes) < 2 Then mes = "0" & mes
		ano = Year(Now)


		If not IsNull(inicio) Then
			data_ini 	= Replace(Left(Inicio,10),"/",".")
    	    hora_ini 	= Mid(Inicio,12,16)
		Else
			'==================================================
			data_ini = dia & "." & mes & "." & ano
			hora_ini = "00:00"
			'==================================================
		End If

		If not IsNull(fim) Then
			data_fim 	= Replace(Left(fim,10),"/",".")
    	    hora_fim 	= Mid(fim,12,16)
		Else
			'==================================================
			data_fim = dia & "." & mes & "." & ano
			hora_fim = "23:59"
			'==================================================
		End If
		ativo 		= RS_Listar("ativo")



		SQL_Lotes = 	"Select * From Edicoes_Lote Where ID_Edicao = " & id & " AND Ativo = 1 Order by Data_Fim ASC"

		Set RS_Lotes = Server.CreateObject("ADODB.Recordset")
		RS_Lotes.Open SQL_Lotes, Conexao

		RS_Listar.Close
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
<script language="javascript" src="/js/jquery.maskedinput-1.2.2.min.js"></script>
<script language="javascript" src="/js/colorpicker/colorpicker.js"></script>
<script language="javascript" src="/js/Calendario/calendar.js"></script>
</head>


<script language="javascript">
$(document).ready(function(){
	$('#hora_ini').mask("99:99",{placeholder:"_"});
	$('#hora_fim').mask("99:99",{placeholder:"_"});
	$('#aviso').hide();
	<%
	If msg = "" AND Session("admin_msg") <> "" Then msg = Session("admin_msg")
	Select Case msg
		Case "pag_proibida"
			Session("admin_msg") = ""
		%>
		$('#aviso_conteudo').html('Página não permitida !');
		$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000);
		<%
	End Select
	%>

});

function Enviar() {
	var erros = 0;
	$('#cad select:enabled').each(function(i) {
		// Se não for obrigatório
		switch (this.id) {
			default:
				erros += verificar(this.id, false);
				break;
		}
	});
	$('#cad input:enabled').each(function(i) {
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
		$('#aviso').fadeIn('slow').animate({opacity: '+=0'}, 3000);
	}
}
function voltar() {
	confirmacao = confirm("Os dados não foram salvos, deseja sair ?")
	if (confirmacao) {
		document.location = 'default.asp';
	}
}

function changeLote(formNow,action){
	$('#acao' + formNow).val(action);
	//alert('cad-lote' + formNow);

	var message = "Tem certeza que deseja remover o lote?";

	if(action == "updateLote"){
		message = "Tem certeza que deseja alterar o lote?"
	}

	confirmacao = confirm(message)
	if (confirmacao) {
		document.getElementById('cad-lote' + formNow).submit();
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
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">Evento ID: <%=id%></span></td>
        <td width="100" align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;"><a href="javascript:voltar();"><img src="/admin/images/bt_voltar_admin.gif" width="100" height="48" border="0" /></a></td>
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
            <input type="hidden" id="acao" name="acao" value="upd_edicao">
            <input type="hidden" id="id" name="id" value="<%=id%>">
              <tr>
                <td height="30" colspan="2" align="center"><span style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">Atualizar Evento</span></td>
                </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Evento:</td>
                <td class="titulo_noticias_home">
				<select id="evento" name="evento" class="admin_txtfield_login">
                <option value="-">-- Selecione --</option>
                <%
				If not RS_Eventos.BOF or not RS_Eventos.EOF Then
					While not RS_Eventos.EOF
						selecionado = ""
						If Cstr(id_evento) = Cstr(RS_Eventos("ID_Evento")) Then selecionado = " selected "
						%><option value="<%=RS_Eventos("ID_Evento")%>" <%=selecionado%>><%=RS_Eventos("Nome_PTB")%></option><%
						RS_Eventos.MoveNext
					Wend
					RS_Eventos.Close
				End If
				%>
                </select>
                </td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_bts">Ano da Edição</td>
                <td class="titulo_noticias_home">
					<select id="ano" name="ano" class="admin_txtfield_login">
						<option value="-">-- Selecione --</option>
						<%
						ano_hj = Year(Now())
						For i = ano_hj To ano_hj + 2
							selecionado = ""
							If Cstr(ano) = Cstr(i) Then selecionado = " selected "
							%><option value="<%=i%>" <%=selecionado%>><%=i%></option><%
						Next
						%>
					</select>
                </td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">In&iacute;cio da Feira</td>
                <td class="t_arial fs11px bold c_vermelho">
                  <input name="data_ini" id="data_ini" type="text" size="12" class="admin_txtfield_login" readonly value="<%=data_ini%>">
                  <img src="/admin/images/img_calendario.gif" width="28" height="24" border="0" align="absmiddle" onClick="displayCalendar(document.forms[0].data_ini,'dd.mm.yyyy',this);" class="bt_aba" id="img_calendario1" />
                  às <input name="hora_ini" id="hora_ini" type="text" size="6" class="admin_txtfield_login" value="<%=hora_ini%>">
                </td>
              </tr>
              <tr>
                <td height="30" class="titulo_menu_site_bts">Término da Feira</td>
                <td class="t_arial fs11px bold c_vermelho">
                  <input name="data_fim" id="data_fim" type="text" size="12" class="admin_txtfield_login" readonly value="<%=data_fim%>">
                  <img src="/admin/images/img_calendario.gif" width="28" height="24" border="0" align="absmiddle" onClick="displayCalendar(document.forms[0].data_fim,'dd.mm.yyyy',this);" class="bt_aba" id="img_calendario1" />
                  às <input name="hora_fim" id="hora_fim" type="text" size="6" class="admin_txtfield_login" value="<%=hora_fim%>">
                </td>
              </tr>
              <tr>
                <td width="260" height="30" class="titulo_menu_site_bts">Evento Disponível</td>
                <td class="titulo_noticias_home">
                  <select id="ativo" name="ativo" class="admin_txtfield_login">
                    <option value="1" <% If ativo = "1" OR ativo = true Then %> selected <% End If %> >Sim</option>
                    <option value="0" <% If ativo = "0" OR ativo = false Then %> selected <% End If %> >Não</option>
                  </select>
                </td>
              </tr>

              <tr>
                <td width="260" height="30" class="titulo_menu_site_tec">&nbsp;</td>
                <td class="titulo_noticias_home">
                  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:left;"></div>
                  <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco cursor" onClick="Enviar();">Atualizar Evento</div>
                  <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:right; width:15px; height:17px; float:left;"></div>
                  </td>
                </tr>
            </form>

			 <tr>
				<td height="30" colspan="2" class="titulo_menu_site_bts"> <hr /> </td>
			 </tr>

			 <tr>
				<td width="510" height="30" colspan="2" class="titulo_menu_site_bts">Lotes</td>
			 </tr>

			 <tr>
				<td height="30" colspan="2" class="titulo_menu_site_bts"> <hr /> </td>
			 </tr>

			<%

			iNow = 1

			If not RS_Lotes.BOF or not RS_Lotes.EOF Then
				While not RS_Lotes.EOF

				raw_hora_ini = RS_Lotes("Data_Inicio")
				raw_hora_fim = RS_Lotes("Data_Fim")


				hora_ini_n 	= Replace(Left(raw_hora_ini,10),"/",".")
				hora_fim_n 	= Replace(Left(raw_hora_fim,10),"/",".")

				%>

				<form name="cad-lote<%=iNow%>" id="cad-lote<%=iNow%>" method="post" action="metodos.asp">
					<input type="hidden" class="id" name="id" value="<%=id%>">
					<input type="hidden" class="id-lote" name="id-lote" value="<%=RS_Lotes("ID_Lote_Edicao")%>">
					<input type="hidden" id="acao<%=iNow%>" class="acao" name="acao" value="">

				<tr>
					<td width="510" height="30" colspan="2" class="titulo_menu_site_bts">Lote: <%=iNow%></td>
				</tr>

				<tr>
					<td colspan="2" width="510">
						<table width="510">
							<tr>
								<td height="30" class="titulo_menu_site_bts">Série</td>
								<td class="t_arial fs11px bold c_vermelho">
								  <input name="serie_lote" id="serie_lote<%=iNow%>" type="text" size="12" class="admin_txtfield_login" value="<%=RS_Lotes("Nome")%>">
								</td>
							 </tr>
							<tr>
								<td height="30" class="titulo_menu_site_bts">Valor do Lote</td>
								<td class="t_arial fs11px bold c_vermelho">
								  R$ <input name="val_lote" id="val_lote<%=iNow%>" type="text" size="6" class="admin_txtfield_login" value="<%=RS_Lotes("Valor")%>">
								</td>
							 </tr>

							 <tr>
								<td height="30" class="titulo_menu_site_bts">Início do Lote</td>
								<td class="t_arial fs11px bold c_vermelho">
								  <input name="data_ini_lote" id="data_ini_lote<%=iNow%>" type="text" size="12" class="admin_txtfield_login" readonly value="<%=hora_ini_n%>">
								  <img src="/admin/images/img_calendario.gif" id="cal-ini-lote<%=iNow%>" width="28" height="24" border="0" align="absmiddle" onClick="displayCalendar(document.getElementById('data_ini_lote<%=iNow%>'),'dd.mm.yyyy',this);" class="bt_aba" id="img_calendario1" />
								</td>
							 </tr>

							 <tr>
								<td height="30" class="titulo_menu_site_bts">Término do Lote</td>
								<td class="t_arial fs11px bold c_vermelho">
								  <input name="data_fim_lote" id="data_fim_lote<%=iNow%>" type="text" size="12" class="admin_txtfield_login" readonly value="<%=hora_fim_n%>">
								  <img src="/admin/images/img_calendario.gif" id="cal-fim-lote" width="28" height="24" border="0" align="absmiddle" onClick="displayCalendar(document.getElementById('data_fim_lote<%=iNow%>'),'dd.mm.yyyy',this);" class="bt_aba" id="img_calendario1" />
								</td>
							 </tr>

							 <tr>
								<td height="30" class="titulo_menu_site_bts">

								</td>
								<td>
									<div class="box-bt">
										<div class="bt-a bt-alterar-lote" onclick="changeLote('<%=iNow%>','updateLote')">Alterar</div>
										<div class="bt-b bt-remover-lote" onclick="changeLote('<%=iNow%>','removeLote')">Remover</div>
									</div>
								</td>
							 </tr>

							 <tr>
								<td height="30" colspan="2" class="titulo_menu_site_bts"> <hr /> </td>
							 </tr>
						</table>
					</td>
				</tr>

				</form>

				<%
				RS_Lotes.MoveNext
				iNow = iNow + 1
				Wend
				RS_Lotes.Close
			End If
			%>



			 <tr>
				<td height="30" class="titulo_menu_site_bts">

				</td>
				<td>
					<div id="inc-lote" class="bt-a">Incluir Lote</div>
				</td>
			 </tr>

			<form id="cad-lote" name="cad-lote" method="post" action="metodos.asp">
				<input type="hidden" class="id" name="id" value="<%=id%>">
				<input type="hidden" class="acao" name="acao" value="insertLote">

			<tr>
				<td colspan="2" width="510">
					<table id="new-lote" width="510">

						<tr>
							<td width="510" height="30" colspan="2" class="titulo_menu_site_bts">Novo Lote</td>
						</tr>

						<tr>
							<td height="30" class="titulo_menu_site_bts">Série</td>
							<td class="t_arial fs11px bold c_vermelho">
							  <input name="serie_lote" id="serie_lote" type="text" size="12" class="admin_txtfield_login" value="">
							</td>
						 </tr>

						<tr>
							<td height="30" class="titulo_menu_site_bts">Valor do Lote</td>
							<td class="t_arial fs11px bold c_vermelho">
							  R$ <input name="val_lote" id="val_lote" type="text" size="6" class="admin_txtfield_login" value="">
							</td>
						 </tr>

						 <tr>
							<td height="30" class="titulo_menu_site_bts">Início do Lote</td>
							<td class="t_arial fs11px bold c_vermelho">
							  <input name="data_ini_lote" id="data_ini_lote" type="text" size="12" class="admin_txtfield_login" readonly value="">
							  <img src="/admin/images/img_calendario.gif" id="cal-ini-lote" width="28" height="24" border="0" align="absmiddle" onClick="displayCalendar(document.getElementById('cad-lote').data_ini_lote,'dd.mm.yyyy',this);" class="bt_aba" id="img_calendario1" />
							</td>
						 </tr>

						 <tr>
							<td height="30" class="titulo_menu_site_bts">Término do Lote</td>
							<td class="t_arial fs11px bold c_vermelho">
							  <input name="data_fim_lote" id="data_fim_lote" type="text" size="12" class="admin_txtfield_login" readonly value="">
							  <img src="/admin/images/img_calendario.gif" id="cal-fim-lote" width="28" height="24" border="0" align="absmiddle" onClick="displayCalendar(document.getElementById('cad-lote').data_fim_lote,'dd.mm.yyyy',this);" class="bt_aba" id="img_calendario1" />
							</td>
						 </tr>

						 <tr>
							<td height="30" class="titulo_menu_site_bts">

							</td>
							<td>
								<div class="bt-a" onclick="valida_lote()">Salvar</div>
							</td>
						 </tr>

						 <tr>
							<td height="30" colspan="2" class="titulo_menu_site_bts"> <hr /> </td>
						 </tr>
					</table>
				</td>
			</tr>


			</form>


          </table></td>
          <td background="/admin/images/caixa/dir.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
        </tr>
        <tr>
          <td><img src="/admin/images/caixa/inf_esq.gif" width="14" height="12" /></td>
          <td background="/admin/images/caixa/inf_centro.gif"><img src="/admin/images/caixa/spacer.gif" width="14" height="12" /></td>
          <td><img src="/admin/images/caixa/inf_dir.gif" width="14" height="12" /></td>
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
