<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/includes/texto_caixaAltaBaixa.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Credenciamento <%=Year(Now())%> - BTS Informa</title>
<link href="/css/base_forms.css" rel="stylesheet" type="text/css" />
<link href="/css/estilos.css" rel="stylesheet" type="text/css">
<link href="/css/jquery.alerts.css" rel="stylesheet" type="text/css">
<link href="/css/checkbox.css" rel="stylesheet" type="text/css">
<link rel="stylesheet" href="/css/colorbox.css" />
<script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
<script language="javascript" src="/js/jquery.maskedinput-1.2.2.min.js"></script>
<script language="javascript" src="/js/jquery-ui-1.8.7.core_eff-slide.js"></script>
<script language="javascript" src="/js/jquery.alerts.js"></script>
<script language="javascript" src="/js/jquery.screwdefaultbuttons.js"></script>
<script language="javascript" src="/js/jquery.colorbox.js"></script>
<script language="javascript" src="/js/validar_forms.js"></script>
<script language="javascript" src="/js/funcoes_gerais.js"></script>
<script language="javascript" src="/js/tipos.js"></script>
<!-- Script desta página -->
<script language="javascript" src="default.js" charset="utf-8"></script>
<!-- Script desta página FIM -->
<%
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================

	'response.write session("teste_paypal'")

				  'response.write nvpstr
If Session("cliente_edicao") = "" OR Session("cliente_idioma") = ""  or Session("cliente_visitante") = "" Then
  response.Redirect("http://www.mbxeventos.net/aol3abf2016/")
End If

ID_Edicao               = Session("cliente_edicao")
Idioma                  = Session("cliente_idioma")
ID_TP_Credenciamento    = Session("cliente_tipo")
TP_Formulario           = Session("cliente_formulario")
IRC                     = Session("cliente_cadastro")
ID_Empresa              = Session("cliente_empresa")
ID_Visitante            = Session("cliente_visitante")
Nome_Visitante          = Session("cliente_nome")
CPF_Visitante           = Session("cliente_cpf")

	' Verifica Idioma a ser apresentado
	Select Case (Idioma)
		Case "1"
			SgIdioma = "PTB"
		Case "2"
			SgIdioma = "ESP"
		Case "3"
			SgIdioma = "ENG"
		Case Else
			SgIdioma = "PTB"
	End Select

	Pagina_ID 	= 2

	SQL_Textos	=	" Select " &_
					"	ID_Texto, " &_
					"	ID_Tipo, " &_
					"	Identificacao, " &_
					"	Texto, " &_
					"	URL_Imagem " &_
					" From Paginas_Textos " &_
					" Where  " &_
					"	ID_Idioma = " & idioma & " " &_
					"	AND ID_Pagina = " & Pagina_ID & " " &_
					" Order By Ordem "
	Set RS_Textos = Server.CreateObject("ADODB.Recordset")
	RS_Textos.Open SQL_Textos, Conexao

	If not RS_Textos.BOF or not RS_Textos.EOF Then
		total_registros = 0
		While not RS_Textos.EOF
			total_registros = total_registros + 1
			RS_Textos.MoveNext
		Wend
		RS_Textos.MoveFirst
		ReDim textos_array(total_registros-1)
		n = 0
		While not RS_Textos.EOF
			id = RS_Textos("id_texto")
			ident = RS_Textos("identificacao")
			texto = RS_Textos("texto")
			url_img = RS_Textos("url_imagem")
			textos_array(n) = Array(id, ident, texto, url_img)
			n = n + 1
			RS_Textos.MoveNext
		Wend
		RS_Textos.Close
	End If


'	For i = Lbound(textos_array) to Ubound(textos_array)
'		response.write("[ i: " & i & " ] [ ident: " & textos_array(i)(1) & " ]  [ txt: " & textos_array(i)(2) & " ]  [ img: " & textos_array(i)(3) & " ]<br>")
'	Next
'===========================================================
%>
<% If Limpar_texto(Request("teste")) = "s" Then %>
	<!--#include virtual="/includes/exibir_array.asp"-->
<% End IF

	' Select IMG Faixa
	SQL_Img_Faixa 	=	"Select " &_
						"	Img_Faixa " &_
						"From Tipo_Credenciamento " &_
						"Where ID_Tipo_Credenciamento = " & ID_TP_Credenciamento
'	response.write(SQL_Img_Faixa & "<br>")
	Set RS_Img_Faixa = Server.CreateObject("ADODB.Recordset")
	RS_Img_Faixa.CursorType = 0
	RS_Img_Faixa.LockType = 1
	RS_Img_Faixa.Open SQL_Img_Faixa, Conexao
		img_faixa = RS_Img_Faixa("img_faixa")
	RS_Img_Faixa.Close

	' Faixa TOPO
	SQL_Faixa	= 	"Select " &_
					"	Cor, " &_
					"	Logo_Negativo, " &_
					"	Faixa_Fundo " &_
					"From Edicoes_configuracao " &_
					"Where  " &_
					"	ID_Edicao = " & ID_Edicao
	Set RS_Faixa = Server.CreateObject("ADODB.Recordset")
	RS_Faixa.CursorType = 0
	RS_Faixa.LockType = 1
	RS_Faixa.Open SQL_Faixa, Conexao

		faixa_cor	= RS_Faixa("cor")
		faixa_logo	= RS_Faixa("logo_negativo")
		faixa_fundo	= RS_Faixa("Faixa_Fundo")
	RS_Faixa.Close

	' Select de Eventos
	SQL_Evento	=	"SELECT " &_
					"	Nome_" & SgIdioma & " AS Evento, " &_
					"	Ano " &_
					"FROM Eventos as E " &_
					"INNER JOIN" &_
					"	Eventos_Edicoes as EE " &_
					"ON EE.ID_Evento = E.ID_Evento " &_
					"WHERE " &_
					"	E.Ativo = 1 " &_
					"	AND EE.ID_Edicao = " & ID_Edicao

	Set RS_Evento = Server.CreateObject("ADODB.Recordset")
	RS_Evento.CursorType = 0
	RS_Evento.LockType = 1
	RS_Evento.Open SQL_Evento, Conexao

	Evento = RS_Evento("Evento") & " " & RS_Evento("Ano")
	Rs_Evento.Close





	'Após o comprovante, caso seja realizado, enviar o email de comprovante
	Session("cliente_enviar_email") = 0

%>
<script language="javascript">
var idioma_atual 	= '<%=Session("cliente_idioma")%>';
var select       	= '<%=textos_array(36)(2)%>';
var cor_fundo 	 	= '<%=faixa_cor%>';
var tp_formulario 	= '';

$(document).ready(function(){
	var erro = '<%=erro%>';
	switch (erro) {
		case '1':
			$('#txt_topo').html('Campos Vazios');
			$('#aviso_topo').show();
			break;
		case '2':
			$('#txt_topo').html('Dados incorretos');
			$('#aviso_topo').show();
			break;
	}
	$(".ajax").colorbox();
});
</script>
</head>

<body>
<div style="width: 100%; position: absolute; height:750px; float:left; z-index:100; background:#CCC; display:none" id="loading" class="transparent">
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center"><img src="/img/geral/ico_ajax-loader.gif" style="opacity:100"></td>
  </tr>
</table>
</div>
<!--#include virtual="/includes/cabecalho.asp"-->
<div style="width: 100%; position: absolute; left:0px; float:left; z-index:10; height: 115px;" id="faixa_selecionada">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="33%" align="center" height="45">
    <!-- Faixa Lateral -->
    	<div style="background:url(/img/geral/faixa_fundo_esq.gif); height:45px; width:100%; margin-top:50px;"></div>
    <!-- Faixa Lateral -->
    </td>
    <td width="870" align="center">
    <!-- Faixa -->
        <table width="870" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr>
            <td height="50">&nbsp;</td>
          </tr>
          <tr>
            <td>
                <table width="870" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="189" height="45" background="/img/geral/faixa_fundo_esqs.gif"><img id="img_faixa_esq" src="/img/geral/tipos/Faixa_Tickets.gif" width="189" height="45" /></td>
                    <td id="img_fundo_selecionado" height="45" background="<%=faixa_fundo%>" class="atencao_13px cor_branco">
                    	<div id="txt_1" style="padding-left:20px; float:left; line-height:40px;" align="left"><!--Preencha os campos abaixo--></div>
                        <div style="float:right;" align="right"><img id="img_logo_selecionado" src="<%=faixa_logo%>" hspace="10"></div>
                    </td>
                  </tr>
                </table>
            </td>
          </tr>
      </table>
    <!-- Faixa -->
    </td>
    <td width="33%" align="center" valign="top">
    <!-- Faixa Lateral -->
    	<div style="background:url(<%=faixa_fundo%>); height:45px; width:100%; margin-top:50px;" id="faixa_dir"></div>
    <!-- Faixa Lateral	 -->
    </td>
  </tr>
  <tr>
    <td align="center">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="center" valign="top">&nbsp;</td>
  </tr>
</table>
</div>
<div style="width: 100%; position: absolute; left:0px; float:left; display:;" id="conteudo">
<table width="870" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="547" height="130" colspan="3">&nbsp;</td>
  </tr>
</table>
    <!-- Form Container -->
    <div id="contForm">
    <!-- Form -->

            <!-- Alert error -->
            <div id="aviso_topo" class="fs_12px arial cor_cinza2">
            	<img src="/img/forms/alert-icon.png" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso">
                &nbsp;<span id="txt_topo"><!--Por favor preencher corretamente os itens em destaque !--><%=textos_array(39)(2)%></span>
			</div><br/>
            <!-- End Alert error -->

            <table width="850" border="0" cellpadding="0" cellspacing="0">
                <tr>
                    <td width="800" height="30" bgcolor="#414042" class="arial fs_13px cor_branco" style="padding-left:15px;"><b>Olá</b> <%=Nome_Visitante%></td>
                    <td width="500" height="30" bgcolor="#414042" align="right"><img src="/img/botoes/voltar.gif" width="47" height="15" hspace="5" class="cursor" onClick="link('status.asp');"></td>
                    <td width="50" height="30" bgcolor="#414042" style=" border-left:#ccc 1px solid;" align="right"><img src="/img/botoes/sair.gif" width="47" height="15" class="cursor" onClick="sair();"></td>
                </tr>
            </table>
            <br/>

            	<%
				SQL_Consulta_Pedidos = 	"Select " &_
										"	P.* " &_
										"From " &_
										"	Pedidos As P " &_

										"Where " &_
										"	P.ID_Edicao = '" & Session("cliente_edicao") & "' " &_
										"	And P.ID_Rel_Cadastro = '" & Session("cliente_cadastro") & "' " &_
										"	And P.ID_Visitante = '" & Session("cliente_visitante")  & "' " &_
										"	And P.Status_Pedido = 1"
				'Response.Write(SQL_Consulta_Pedidos)

				Set RS_Consulta_Pedidos = Server.CreateObject("ADODB.Recordset")
				RS_Consulta_Pedidos.Open SQL_Consulta_Pedidos, Conexao, 3, 3

				If Not RS_Consulta_Pedidos.Eof Then

					Tickets 		= True
					Numero_Pedido 	= RS_Consulta_Pedidos("Numero_Pedido")
					'Session("pedido") = RS_Lista_Pedidos("Numero_Pedido")
					ID_Pedido 		= RS_Consulta_Pedidos("ID_Pedido")
					Idioma_Pedido	= RS_Consulta_Pedidos("ID_Idioma")
					Valor_Pedido	= FormatNumber(RS_Consulta_Pedidos("Valor_Pedido"),2)


				Else

					Tickets = False

				End If
				%>

				<!--#Include virtual="/tickets/menu_lateral.asp"-->

	<form action="expresscheckout.asp" method="post" id="Pagamento" name="Pagamento" >

            <fieldset style="float: right; width: 580px; ">
				<%
					SQL_Consulta_Pedidos =	"Select " &_
											"	P.* " &_
											"	,V.Nome_Completo " &_
											"	,V.CPF " &_
											"	,V.Passaporte, V.Email " &_
											"From Pedidos As P " &_
											"Left Join Visitantes As V On V.ID_Visitante = P.ID_Visitante " &_
											"Where " &_
											"	P.ID_Edicao = '" & Session("cliente_edicao") & "'  " &_
											"	And P.ID_Rel_Cadastro = '" & Session("cliente_cadastro") & "'  " &_
											"	And P.ID_Visitante = '" & Session("cliente_visitante")  & "'  " &_
											"	And P.Status_Pedido = 1 "

										'Response.Write(SQL_Consulta_Pedidos)

                    Set RS_Consulta_Pedidos = Server.CreateObject("ADODB.Recordset")
                    RS_Consulta_Pedidos.Open SQL_Consulta_Pedidos, Conexao, 3, 3
                    session("Numero_Pedido") =""
                    If Not RS_Consulta_Pedidos.Eof Then

                        Tickets 		= True
                        Numero_Pedido 	= RS_Consulta_Pedidos("Numero_Pedido")
						session("Numero_Pedido") = Numero_Pedido
                        ID_Pedido 		= RS_Consulta_Pedidos("ID_Pedido")
                        Idioma_Pedido	= RS_Consulta_Pedidos("ID_Idioma")
                        Valor_Pedido	= FormatNumber(RS_Consulta_Pedidos("Valor_Pedido"),2)
						ID_Visitante	= RS_Consulta_Pedidos("ID_Visitante")
						Nome_Completo	= RS_Consulta_Pedidos("Nome_Completo")
						CPF				= RS_Consulta_Pedidos("CPF")
						Passaporte		= RS_Consulta_Pedidos("Email")

						Cliente = "<strong>" & Nome_Completo & "</strong><br>"
						If Len(Trim(CPF)) = 11 Then
						Cliente = Cliente & "<strong>CPF</strong>: " & CPF
						Else
						Cliente = Cliente & "<strong>E-mail</strong>: " & Passaporte
						End If

                    End If
					Valor_PedidoEnviado = replace(Valor_Pedido,".00","")
					Valor_PedidoEnviado = replace(Valor_Pedido,",00","")

                %>
            	<legend>Detalhes do Pedido</legend>
				<div id="parcAssis" class="div_parceria" style="width: 580px; padding-bottom:15px;">
                	<input type="hidden" name="ValorDocumento" id="ValorDocumento" value="<%=Valor_PedidoEnviado %>"/>
                    <input type="hidden" name="NumeroDocumento" id="NumeroDocumento" value="<%=Numero_Pedido%>"/>
                    <input type="hidden" name="Moeda" id="Moeda" value="BRL"/>
                    <input type="hidden" name="PreAutorizacao" id="PreAutorizacao" value="2"/>
                    <input type="hidden" name="QuantidadeParcelas" id="QuantidadeParcelas" value="1"/>
                    <input type="hidden" name="ParametrosCliente" id="ParametrosCliente" value="<%=Cliente%>"/>


                        <div style="width: 575px; padding: 5px 0 5px 5px">Pedido nº: &nbsp;<font style="font-size: 16px;"><%=Numero_Pedido%></font></div>

                        <table cellpadding="0" cellspacing="0" width="100%">
                        	<tr>
                            	<td style="width: 575px;" colspan="3">

                                	<table cellpadding="0" cellspacing="3" width="100%" style="padding: 10px 0 10px;">
                                    	<tr>
                                        	<td bgcolor="CCCCCC" style="padding: 5px; width: 100px; font-weight: 100">ID Usuário:</td>
                                            <td bgcolor="f1f0f0" style="padding: 5px;"><%=ID_Visitante%></td>
                                        </tr>
                                    	<tr>
                                        	<td bgcolor="CCCCCC" style="padding: 5px; width: 100px; font-weight: 100">Nome Completo:</td>
                                            <td bgcolor="f1f0f0" style="padding: 5px;"><%=Nome_Completo%></td>
                                        </tr>
                                    	<tr>
                                        	<td bgcolor="CCCCCC" style="padding: 5px; width: 100px; font-weight: 100"><%If Len(Trim(CPF)) =11   Then%>CPF<%Else%>E-mail<%End If%>:</td>
                                            <td bgcolor="f1f0f0" style="padding: 5px;"><%If Len(Trim(CPF)) =11 Then Response.Write(CPF) Else Response.Write(Passaporte)%></td>
                                        </tr>
                                    </table>

                                </td>
                            </tr>
                        	<tr>
                            	<td style="padding: 5px; width: 575px; font-weight: 100" colspan="3">Segue abaixo a lista completa de Visitantes em seu PEDIDO:</td>
                            </tr>

                        	<tr bgcolor="CCCCCC">
                            	<td style="padding: 5px; width: 375px;">NOME COMPLETO</td>
                                <td style="padding: 5px; width: 100px;">TIPO</td>
                                <td style="padding: 5px; width: 100px;">DOCUMENTO</td>
                            </tr>
                        <%
                        SQL_Carrinho = 	"Select " &_
                                        "	C.ID_Carrinho,  " &_
                                        "	C.ID_Visitante,  " &_
                                        "	C.ID_Pedido,  " &_
                                        "	C.ID_Rel_Cadastro, " &_
                                        "	P.Status_Pedido, " &_
                                        "	V.Nome_Completo, " &_
										"	V.CPF, " &_
										"	V.Passaporte, V.EMAIL " &_
                                        "From  Pedidos_Carrinho  As C " &_
                                        "Inner Join Visitantes As V On V.ID_Visitante = C.ID_Visitante " &_
                                        "Inner Join Pedidos As P On P.ID_Pedido = C.ID_Pedido " &_
                                        "Where " &_
										"	C.Cancelado = 0 " &_
                                        "	AND C.ID_Pedido = " & ID_Pedido
                        Set RS_Carrinho = Server.CreateObject("ADODB.Recordset")
                        RS_Carrinho.Open SQL_Carrinho, Conexao, 3, 3

                        Primeiro = 0
						Z = True
                        session("finaliza")=""
                        If Not RS_Carrinho.Eof Then

                            While Not RS_Carrinho.Eof

							If Z = True Then
								Cor_Fundo = "e5e5e5"
							Else
								Cor_Fundo = "efefef"
							End If

							Z = Not Z
                        %>

                        	<tr bgcolor="<%=Cor_Fundo%>" style="padding: 5px; font-weight: 100">
                            	<td style="padding: 5px; width: 375px;"><%=RS_Carrinho("Nome_Completo")%></td>
                                <td style="padding: 5px; width: 100px;"><%If Len(Trim(RS_Carrinho("CPF"))) = 11 Then%>CPF<%Else%>E-mail<%End If%></td>
                                <td style="padding: 5px; width: 100px;"><%If Len(Trim(RS_Carrinho("CPF"))) = 11 Then Response.Write(RS_Carrinho("CPF")) Else Response.Write(RS_Carrinho("Email"))%></td>
                            </tr>

                        <%
						valor = RS_Carrinho("CPF")
						If Len(Trim(RS_Carrinho("CPF"))) <> 11 then valor = RS_Carrinho("email")
						if session("finaliza") = "" then
								session("finaliza") = valor
						else
						session("finaliza") = session("finaliza") & "," & valor
						end if
                            RS_Carrinho.MoveNext
                            Wend
                        End If
                        %>
                        </table>

                        <div style="width: 575px;  border-bottom: 1px dotted #999; font-size: 14px; padding: 5px 0 5px 5px; background: #CCC">
                            <font style="font-weight: 100;">Valor Total: &nbsp;</font>
                            <strong><%If Cint(Idioma_Pedido) = 1 Then Response.Write("R$") Else Response.Write("$")%>&nbsp;<%=Valor_Pedido%></strong>
                        </div>
                        <div style="width: 575px; height: 30px; margin-top: 15px;">
							<a class='ajax' href="saiba-mais_04.jpg"><div>Por que comprar com PayPal?</div></a>
                            <a href="/tickets/novo_pedido.asp"><div class="bt_alterar_compra" style="float: left">Alterar Compra</div></a>
                            <a href="#finalizar_pedido" onclick="finaliza_paypal();"><div class="continuar" style="float: right">Continuar</div></a>
                        </div>
                	</div>
            </fieldset>
            <br/>

            <!-- Alert error -->
            <div id="aviso" class="fs_12px arial cor_cinza2" style="display:inline-table; margin-top:15px;">
            	<img src="/img/forms/alert-icon.png" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso">
                &nbsp;<span id="txt"><!--Por favor preencher corretamente os itens em destaque !--><%=textos_array(39)(2)%></span>
			</div>
            <!-- End Alert error -->
        </form>
        <!-- Form End -->
	</div>
    <!-- End Form Container -->
<table width="870" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="547" height="50" colspan="3">&nbsp;</td>
  </tr>
</table>
</div>
<div style="width: 100%; position: absolute;float:left; display:none; z-index:100" id="loading">
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center"><img src="/img/geral/ico_ajax-loader.gif" style="opacity:100"></td>
  </tr>
</table>
</div>
</body>
</html>
<script type="text/javascript">

function finaliza_paypal(){


			jQuery('<div class="sa_payPal_overlay" style="visibility:visible;position:fixed; width:100%; height:100%; filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=1, StartColorStr=\'#88ffffff\', EndColorStr=\'#88ffffff\'); background: rgba(255,255,255,0.8); top:0; left:0; z-index: 999999;"><div style=" background: #FFF; background-image: linear-gradient(top, #FFFFFF 45%, #E9ECEF 80%);background-image: -o-linear-gradient(top, #FFFFFF 45%, #E9ECEF 80%);background-image: -moz-linear-gradient(top, #FFFFFF 45%, #E9ECEF 80%);background-image: -webkit-linear-gradient(top, #FFFFFF 45%, #E9ECEF 80%);background-image: -ms-linear-gradient(top, #FFFFFF 45%, #E9ECEF 80%);background-image: -webkit-gradient(linear, left top,left bottom,color-stop(0.45, #FFFFFF),color-stop(0.8, #E9ECEF));display: block;margin: auto;position: fixed; margin-left:-220px; left:47%;top: 33%;text-align: center;color: #2F6395;font-family: Arial;padding: 15px;font-size: 15px;font-weight: bold;width: 530px;-webkit-box-shadow: 3px 2px 13px rgba(50, 50, 49, 0.25);box-shadow: rgba(0, 0, 0, 0.2) 0px 0px 0px 5px;border: 1px solid #CFCFCF;border-radius: 6px;"><img style="display:block;margin:0 auto 10px" src="https://www.paypalobjects.com/en_US/i/icon/icon_animated_prog_dkgy_42wx42h.gif"><h2 style="color:inherit !important;background:none !important;border:none !important;font-size:23px !important;text-decoration:none !important;text-transform:none !important;font-family: Arial !important;">Aguarde alguns segundos.</h2> <p style="font-size:13px; margin-top:13px; color: #003171; font-weight:400">Você está sendo redirecionado para um ambiente seguro para finalizar sua compra.</p><div style="margin:20px auto 0;"><img src="https://www.paypal-brasil.com.br/logocenter/util/img/logo_paypal.png"/></div></div></div>').appendTo('body');

			$('#Pagamento').submit();
}




</script>
<%
Conexao.Close
%>
