<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/includes/texto_caixaAltaBaixa.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Credenciamento <%=Year(Now())%> - BTS Informa</title>

<!-- Script desta página FIM -->
<%
'response.end
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================



'For Each item In Request.Form
'	Response.Write "" & item & " - Value: " & Request.Form(item) & "<BR />"
'Next


If Session("cliente_edicao") = "" OR Session("cliente_idioma") = "" OR Session("cliente_logado") = "" or Session("cliente_visitante") = "" Then
    response.Redirect("http://www.mbxeventos.net/AOLABF2017/")
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
<% If Limpar_Texto(Request("teste")) = "s" Then %>
	<!--#include virtual="/includes/exibir_array.asp"-->
<% End IF

	' Select IMG Faixa
	SQL_Img_Faixa 	=	"Select " &_
						"	Img_Faixa " &_
						"From Tipo_Credenciamento " &_
						"Where ID_Tipo_Credenciamento = " & ID_TP_Credenciamento
	'response.write(SQL_Img_Faixa & "<br>")
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

%>
<link href="/css/base_forms.css" rel="stylesheet" type="text/css" />
<link href="/css/estilos.css" rel="stylesheet" type="text/css">
<link href="/css/jquery.alerts.css" rel="stylesheet" type="text/css">
<link href="/css/checkbox.css" rel="stylesheet" type="text/css">

<script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
<script language="javascript" src="/js/jquery-ui-1.8.7.core_eff-slide.js"></script>
<script language="javascript" src="/js/jquery.alerts.js"></script>
<script language="javascript" src="/js/validar_forms.js"></script>
<script language="javascript" src="/js/funcoes_gerais.js"></script>
<script language="javascript" src="/js/jquery.translate.js"></script>

<!-- Script desta página -->
<script language="javascript" src="default.js" charset="utf-8"></script>

<script language="javascript">
$(function() {

  var t = {
    text1: {
	  pt: "Olá",
      en: "Hello",
	  es: "Hola"
    },
	text2: {
	  pt: "Menu",
      en: "Menu",
	  es: "Menú"
    },
	text3: {
	  pt: "Meus Pedidos",
      en: "My purchase orders",
	  es: "Mis pedidos"
    },
	text4: {
	  pt: "Continuar pedido",
      en: "Continue with the purchase order",
	  es: "Continuar pedido"
    },
	text5: {
	  pt: "Suporte: visitante.abf@informa.com",
      en: "Support: visitante.abf@informa.com",
	  es: "Asistencia: visitante.abf@informa.com"
    },
	text6: {
	  pt: "Suporte:",
      en: "Support:",
	  es: "Asistencia:"
    },
	text7: {
	  pt: "Pedido n°:",
      en: "Purchase Order number:",
	  es: "Pedido número:"
    },
	text8: {
	  pt: "Pessoas em meu pedido:",
      en: "People in my purchase order:",
	  es: "Personas en mi pedido:"
    },
	text9: {
	  pt: "Compre ingresso para outras pessoas",
      en: "Buy tickets for other people",
	  es: "Comprar entradas para los demás"
    },
	text10: {
	  pt: "Ainda não há pessoas adicionadas ao seu Pedido!",
      en: "There are no people added to your order!",
	  es: "No hay personas agregado a su petición!"
    },
	text11: {
	  pt: "Utilize a busca abaixo para poder adicionar pessoas ao seu Pedido.",
      en: "Use the search below to add people to your Order.",
	  es: "Utilice el buscador de abajo para poder agregar personas a su solicitud."
    },
	text12: {
	  pt: "Valor Total:",
      en: "Total Amount:",
	  es: "Cantidad Total:"
    },
	text13: {
	  pt: "Atenção",
      en: "Attention",
	  es: "Attention"
    },
	text14: {
	  pt: "Cada ingresso é válido para os 04 dias da feira, podendo ser adquiridos antecipadamente por meio do site do evento ao custo de R$ 60,00 até dia 20 de junho de 2017.",
      en: "Each ticket is valid for the 4 days of the fair, and can be purchased in advance on the event’s website at the cost of BRL$ 60.00 until June 20, 2017.",
	  es: "Cada billete es válido para los 04 días de la feria, que se pueden adquirir con antelación por medio del sitio web del evento con un costo de R$ 60,00 hasta el día 20 de junio de 2017."
    },
	text15: {
	  pt: "Caso deixe para comprar durante a realização do evento, de 21 a 24 de junho de 2017, na bilheteria local ou pelo site custará R$ 70,00.",
      en: "If you prefer to buy the ticket during the event, from June 21 to 24, 2017, at the local ticket office or on the website, the ticket will cost BRL$ 70.00.",
	  es: "Si realiza la compra durante la realización del evento, del 21 al 24 de junio de 2017, en la boletería del local o a través del sitio web costará R$ 70,00."
    },
	text16: {
	  pt: "Atenção",
      en: "Attention",
	  es: "Attention"
    },
	text17: {
	  pt: "Concluir a compra",
      en: "Complete this purchase order",
	  es: "Concluir este pedido"
    },
	text18: {
	  pt: "Se quiser comprar ingressos para outras pessoas, utilize o quadro de busca abaixo. A busca deverá ser feita pelo número do <font style='font-weight: bold'>CPF</font> ou <font style='font-weight: bold'>E-mail</font>, em caso de estrangeiros.<br><font style='font-size: 10px'><em>Obs.: para que o <strong>CPF</strong> ou <strong>E-mail</strong> constem em nossa base de dados, é necessário que estas pessoas já tenham feito seu credenciamento.</em></font>",
      en: "If you want to buy tickets for other people, use the search chart below.  The search shall be done by the <font style='font-weight: bold'>CPF number</font> or <font style='font-weight: bold'>E-mail</font>, in case of foreigners.<br><font style='font-size: 10px'><em>Note: For the <strong>CPF</strong> or <strong>E-mail</strong> to appear in our database, it is necessary for these people to have been already accredited.</em></font>",
	  es: "Si desea comprar los billetes para otras personas, utilice el campo de búsqueda a continuación. La búsqueda se realizar con el número de identificación fiscal  <font style='font-weight: bold'>CPF</font> o <font style='font-weight: bold'>E-mail</font>, en el caso de extranjeros.<br><font style='font-size: 10px'><em>Obs.: para que el número de identificación fiscal <strong>CPF</strong> o <strong>E-mail</strong> consten en nuestra base de datos, es necesario que estas personas ya hayan realizado su acreditación.</em></font>"
    },
	text19: {
	  pt: "BUSCA:",
      en: "SEARCH:",
	  es: "BÚSQUEDA"
    },
	text20: {
	  pt: "Nova Compra",
      en: "New purchase",
	  es: "Nuevo pedido"
    },
	text21: {
	  pt: "Você ainda não possui Pedido para este evento. Clique no botão abaixo para iniciar o seu primeiro Pedido.",
      en: "You do not have Purchase Orders for this event yet.  Click On the button below to start your first Purchase Order.",
	  es: "Todavía no tiene un Pedido para este evento. Haga clic en el botón a continuación para iniciar su primer Pedido."
    },
	text22: {
	  pt: "Comprar meu Ticket Agora",
      en: "Purchase a ticket now",
	  es: "Comprar billete ahora"
    },
	text23: {
	  pt: "Novo Pedido",
      en: "New purchase order ",
	  es: "Nuevo pedido"
    }
	
  };
  var cookieLang = readCookie("lang");
  if(cookieLang == null)
	var _t = $('body').translate({lang: "pt", t: t});
   else
   var _t = $('body').translate({lang: cookieLang, t: t});
   
  var str = _t.g("translate");

  
  $(".lang_selector").click(function(ev) {
    var lang = $(this).attr("data-value");
    _t.lang(lang);
	createCookie("lang",lang,100);
    ev.preventDefault();
  });
});

</script>

<script language="javascript">
var idioma_atual 	= '<%=Session("cliente_idioma")%>';
var select       	= '<%=textos_array(36)(2)%>';
var cor_fundo 	 	= '<%=faixa_cor%>';
var tp_formulario 	= '';

$(document).ready(function(){
	var erro = '<%=Limpar_Texto(Request("erro"))%>';
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
                    <td width="189" height="45" background="/img/geral/faixa_fundo_sesq.gif"><img id="img_faixa_esq" src="/img/geral/tipos/Faixa_Tickets.gif" width="189" height="45"></td>
                    <td id="img_fundo_selecionado" height="45" background="<%=faixa_fundo%>" style="background-repeat:repeat-x; position:relative;" class="atencao_13px cor_branco">
                    	<div id="txt_1" style="padding-left:20px; float:left; line-height:40px;" align="left"></div>
                        <div style="position:absolute; top:-45px; right:0px;" align="right"><img id="img_logo_selecionado" src="<%=faixa_logo%>" hspace="10"></div>
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
            <!-- Alert error -->
            <div id="aviso_topo" class="fs_12px arial cor_cinza2">
            	<img src="/img/forms/alert-icon.png" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso">
                &nbsp;<span id="txt_topo"><!--Por favor preencher corretamente os itens em destaque !--><%=textos_array(39)(2)%></span>
			</div><br/>
            <!-- End Alert error -->

            <table width="850" border="0" cellpadding="0" cellspacing="0">
                <tr>
                    <td width="800" height="30" bgcolor="#414042" class="arial fs_13px cor_branco" style="padding-left:15px;"><b class="trn" data-trn-key="text1">Olá</b> <%=Nome_Visitante%></td>
                    <td width="500" height="30" bgcolor="#414042" align="right"><img src="/img/botoes/voltar.gif" width="47" height="15" hspace="5" class="cursor" onClick="link('status.asp');"></td>
                    <td width="50" height="30" bgcolor="#414042" style=" border-left:#ccc 1px solid;" align="right"><img src="/img/botoes/sair.gif" width="47" height="15" class="cursor" onClick="CloseWindow();"></td>
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
					QtdeFinal 		= RS_Consulta_Pedidos("Quantidade")

				Else

					Tickets = False
					QtdeFinal = 1

				End If

        'Valida se pedido está no lote correto
        SQL_Valor_Ticket = "select top 1 * from Edicoes_lote where " &_
        											"ID_Edicao = '" & Session("cliente_edicao") & "' " &_
        											"and Ativo = 1 and GETDATE() between Data_Inicio and Data_Fim order by Data_fim asc"


        Set RS_Consulta_Pedidos = Server.CreateObject("ADODB.Recordset")
        RS_Consulta_Pedidos.Open SQL_Valor_Ticket, Conexao, 3, 3

        If Not RS_Consulta_Pedidos.Eof Then

        	Valor_Ticket_Atualizado = FormatNumber(RS_Consulta_Pedidos("Valor"),2)

        Else

        	Response.Write("EDICAO NAO CADASTRADA")
        	Response.End

        End If

		'Response.Write QtdeFinal
		'Response.End

		valor_pedido_por_ingresso = FormatNumber(Valor_Pedido / CInt(QtdeFinal),2)

        If Valor_Ticket_Atualizado <> Valor_Pedido And Valor_Pedido <> "" Then

          Valor_Pedido = FormatNumber(Valor_Ticket_Atualizado * CInt(QtdeFinal),2)

		  Valor_Pedido = Replace(Valor_Pedido, ",", ".")

          Valor_Ticket_Atualizado = Replace(Valor_Ticket_Atualizado, ",", ".")

          SQL_Atualiza_Valor_Pedido = 	"Update Pedidos Set " &_
                      "	Valor_Pedido = " & Valor_Pedido  & " " &_
                      "Where ID_Pedido = " & ID_Pedido

          Set RS_Atualiza_Pedido = Conexao.Execute(SQL_Atualiza_Valor_Pedido)

          Lote_Mudou = "O lote virou e o valor do ingresso mudou. Por favor confira o novo valor."

        End If

        'FIM Valida se pedido está no lote correto

				%>

				<!--#Include virtual="/tickets/menu_lateral.asp"-->

                <form id="form_pedidos" name="form_pedidos" onsubmit="return false" action="/tickets/pedido.asp" method="post">
                    <fieldset style="float: right; width: 580px;">
                    <input type="hidden" id="aceito" name="aceito" value="1">

                        <%
							'SQL_Cancelado = "Select " &_
							'				"	Cancelado " &_
							'				"From Pedidos_Carrinho " &_
							'				"Where" &_
							'				"	P.ID_Visitante = '" & Session("cliente_visitante")  & "' " &_
							'				"	And P.ID_Pedido = '" & ID_Pedido & "'"

							'Response.Write(SQL_Cancelado)

							'Set RS_Cancelado = Server.CreateObject("ADODB.Recordset")
							'RS_Cancelado.Open SQL_Cancelado, Conexao, 3, 3

							'Cancelado 		= SQL_Cancelado("Cancelado")
							If Tickets = False Then
						%>
                        <legend  class="trn" data-trn-key="text20">Nova Compra</legend>
                        <div id="parcAssis" class="div_parceria" style="width:580px; float: right; margin-top: 10px;">
                            <div style="padding: 10px 0 0; font-weight: 100">
                            	<p class="trn" data-trn-key="text21">Você ainda não possui Pedido para este evento. Clique no botão abaixo para iniciar o seu primeiro Pedido.</p>
                                <a href="#comprar-tickets" onclick="link('termo.asp')"><div class="bt_comprar_ticket trn" style="margin-top: 10px" data-trn-key="text22">Comprar meu Ticket Agora</div></a>
                            </div>
						<%Else


							If Session("Novo_Pedido") = True Then
								Session("Novo_Pedido") = False
							%>
							<script language="javascript">
								<%
								If Session("Possui_Pedido") = True Then
									Session("Possui_Pedido") = False

									Texto_Pedido = "Nova Compra em aberto. <br><br> Texto para quando o visitante já possuir um pedido com outro Visitante. <br><br>Nesta tela, você poderá adicionar Tickets para outras pessoas.<br><br> Para ter mais detalhes sobre seu Compra, clique em <strong>Minhas Compras</strong> no menu lateral."
								Else
									Texto_Pedido = "Seu pedido foi gerado com <strong>Sucesso</strong>!<br><br>Nesta tela, você poderá adicionar Tickets para outras pessoas.<br><br> Para ter mais detalhes sobre sua Compra, clique em <strong>Minhas Compras</strong> no menu lateral."
								End If
								%>

								// Aviso comentado por solicitação da Stefanie 08-03-2013
								//jAlert('<%=Texto_Pedido%>','Novo Pedido');
							</script>
							<%End If%>

							<legend><b class="trn" data-trn-key="text7">Pedido nº:</b> <font style="font-size: 16px"><%=Numero_Pedido%></font></legend>
							<div id="parcAssis" class="div_parceria" style="width:580px; float: right; margin-top: 10px;">


                                    <%
									SQL_Carrinho_Cancelado = 	"Select " &_
																"	C.ID_Carrinho,  " &_
																"	C.ID_Visitante,  " &_
																"	C.ID_Pedido,  " &_
																"	C.ID_Rel_Cadastro, " &_
																"	P.Status_Pedido, " &_
																"	V.Nome_Completo, " &_
																"	C.Cancelado " &_
																"From  Pedidos_Carrinho  As C " &_
																"Inner Join Visitantes As V On V.ID_Visitante = C.ID_Visitante " &_
																"Inner Join Pedidos As P On P.ID_Pedido = C.ID_Pedido " &_
																"Where " &_
																"	C.Cancelado = 1 " &_
																"	AND C.ID_Pedido = " & ID_Pedido


									'Response.Write(SQL_Carrinho)
									Set RS_Carrinho_Cancelado = Server.CreateObject("ADODB.Recordset")
									RS_Carrinho_Cancelado.Open SQL_Carrinho_Cancelado, Conexao, 3, 3

									' Se existirem itens cancelados
									If not RS_Carrinho_Cancelado.BOF or not RS_Carrinho_Cancelado.EOF Then
									%>
                                        <div style="margin-top: 10px; width: 575px; float: left; border-top: 1px dotted #999; font-size: 14px; padding: 5px 0 5px 5px; background:#900; color:white;">
                                            Atenção: as pessoas abaixo foram removidas de seu carrinho.*<br />
                                            <small><i>*&nbsp;Já possuem ingressos adquiridos.</i></small>
                                        </div>
                                	<%
										'Loop nos registros
										While not RS_Carrinho_Cancelado.EOF
											%>
												<div style="width: 570px; float: left; font-size: 12px; border-top: 1px dotted #999; padding: 5px;">
													<div style="float: left; width: 470px; padding: 5px; color:#900;"><i><%=RS_Carrinho_Cancelado("Nome_Completo")%></i></div>
												</div>
											<%
											RS_Carrinho_Cancelado.MoveNext
										Wend
										RS_Carrinho_Cancelado.Close
									End If


								'If Cancelado = 1 Then
									SQL_Carrinho = 	"Select " &_
													"	C.ID_Carrinho,  " &_
													"	C.ID_Visitante,  " &_
													"	C.ID_Pedido,  " &_
													"	C.ID_Rel_Cadastro, " &_
													"	P.Status_Pedido, " &_
													"	V.Nome_Completo, " &_
													"	C.Cancelado " &_
													"From  Pedidos_Carrinho  As C " &_
													"Inner Join Visitantes As V On V.ID_Visitante = C.ID_Visitante " &_
													"Inner Join Pedidos As P On P.ID_Pedido = C.ID_Pedido " &_
													"Where " &_
													"	C.Cancelado = 0 " &_
													"	AND C.ID_Pedido = " & ID_Pedido


									'Response.Write(SQL_Carrinho)
									Set RS_Carrinho = Server.CreateObject("ADODB.Recordset")
									RS_Carrinho.Open SQL_Carrinho, Conexao, 3, 3
								'End if
								Primeiro = 0

								If Not RS_Carrinho.Eof Then

								%>
								<div style="margin-top: 10px; width: 575px; float: left; border-top: 1px dotted #999; font-size: 14px; padding: 5px 0 5px 5px; background: #ffd51f">
                                	<label class="trn" data-trn-key="text8">Pessoas em meu pedido:</label>
                                    <%
									SQL_Carrinho_Usuarios = "Select " &_
															"	Count(*) As Quantidade " &_
															"From  Pedidos_Carrinho  As C " &_
															"Inner Join Visitantes As V On V.ID_Visitante = C.ID_Visitante " &_
															"Inner Join Pedidos As P On P.ID_Pedido = C.ID_Pedido " &_
															"Where " &_
															"	C.ID_Pedido = " & ID_Pedido
									Set RS_Carrinho_Usuarios = Server.CreateObject("ADODB.Recordset")
									RS_Carrinho_Usuarios.Open SQL_Carrinho_Usuarios, Conexao, 3, 3
										If Not RS_Carrinho.Eof Then
											Quantidade = RS_Carrinho_Usuarios("Quantidade")
										End If
									RS_Carrinho_Usuarios.Close

									If Cint(Quantidade) > 1 Then
										%><div style="float: right; width: 70px; display: block;">Remover</div><%
									End If
									%>
                                </div>
								<%

								Erro = False

									While Not RS_Carrinho.Eof

										If Cstr(RS_Carrinho("ID_Visitante")) = Cstr(Session("cliente_visitante")) Then
											Primeiro = 1
											Bt_Excluir = ""
										Else
											Bt_Excluir = "<img src='/img/forms/delete.png' onclick='RemoverVisitante("&RS_Carrinho("ID_Carrinho")&","&ID_Pedido&")' style='float: left; margin-top: -3px; cursor: pointer;' alt='Remover este Visitante do meu Compra!' title='Remover este Visitante do meu Compra!'>"
										End If
									%>
										<div style="width: 570px; float: left; font-size: 12px; border-top: 1px dotted #999; padding: 5px;">
											<div style="float: left; width: 470px; padding: 5px;"><%=RS_Carrinho("Nome_Completo")%></div>
											<div style="float: right; width: 30px; padding: 5px;"><%=Bt_Excluir%></div>
										</div>
									<%
										RS_Carrinho.MoveNext
									Wend
								Else

								Erro = True

								%>
								<div style="margin-top: 10px; width: 575px; float: left; border-top: 1px dotted #999; font-size: 14px; padding: 5px 0 5px 5px; background: #ffd51f">
                                	<label class="trn" data-trn-key="text9">Compre ingresso para outras pessoas</label>
                                </div>

                                <div style="display: none;">
									<div style="width: 570px; height: 150px; float: left; font-size: 12px; border-top: 1px dotted #999; padding: 10px 5px;">
										<label class="trn" data-trn-key="text10">Ainda não há pessoas adicionadas ao seu Pedido!</label> <br><label class="trn" data-trn-key="text11">Utilize a busca abaixo para poder adicionar pessoas ao seu Pedido.</label>
									</div>
								<%
								End If
								RS_Carrinho.Close
								%>

								<div style="width: 575px; float: left; font-size: 14px; border-bottom: 1px dotted #999; font-size: 14px; padding: 5px 0 5px 5px; background: #CCC">
									<font style="font-weight: 100;" class="trn" data-trn-key="text12">Valor Total: &nbsp;</font>
									<strong><%If Cint(Idioma_Pedido) = 1 Then Response.Write("R$") Else Response.Write("$")%>&nbsp;<%=Valor_Pedido%></strong>
								</div>

                <%
                  If Lote_Mudou <> "" Then
                %>
								<!--
								<div style="width: 575px; float: left; font-size: 11px; border-bottom: 1px dotted #999; font-size: 11px; padding: 5px 0 5px 5px; color: red;">
  									<font style="font-weight: 900;">Atenção: &nbsp;</font>
  									<span><%=Lote_Mudou%></span>
  								</div>
								-->
                <%
                  End If
                %>

                                <%If Erro = True Then%>
                                </div>
                                <%End If%>

								<%If Erro = False Then%>

								<div class="" style="font-size: smaller">
									<br>
									<p class="trn" data-trn-key="text13">Atenção</p>
									<span class="trn" data-trn-key="text14">Cada ingresso é válido para os 04 dias da feira, podendo ser adquiridos antecipadamente por meio do site do evento ao custo de R$ 60,00 até dia 20 de junho de 2017. </span>
									<br>
									<span class="trn" data-trn-key="text15">Caso deixe para comprar durante a realização do evento, de 21 a 24 de junho de 2017, na bilheteria local ou pelo site custará R$ 70,00</span>
									</div>

								<table width="575" cellpadding="5" style=" display:none; text-align:left; margin-top:20px; float:left; background: #f5f5f5;">

									<tr>
										<td width="250"><span style="font-size:16px; font-weight:900;" class="trn" data-trn-key="text16">Atencao:</span></td>
										<td></td>
										<td></td>
									</tr>

									<!-- <tr style="font-size:15px;">
										<td><strong>Lote</strong></td>
										<td><strong>Valor</strong></td>
										<td><strong>Data de encerramento do lote</strong></td>
									</tr> -->


									<%

									iNow = 0

									SQL_Lotes = 	"Select * From Edicoes_Lote Where ID_Edicao = " & ID_Edicao & " AND Ativo = 1 Order by Data_Fim ASC"

									Set RS_Lotes = Server.CreateObject("ADODB.Recordset")
									RS_Lotes.Open SQL_Lotes, Conexao


									If not RS_Lotes.BOF or not RS_Lotes.EOF Then
										While not RS_Lotes.EOF

										raw_hora_ini = RS_Lotes("Data_Inicio")
										raw_hora_fim = RS_Lotes("Data_Fim")

										hora_ini_n 	= Replace(Left(raw_hora_ini,10),"/",".")
										hora_fim_n 	= Replace(Left(raw_hora_fim,10),"/",".")

										hora_ini_t = Left(Right(raw_hora_ini,8),5)
										hora_fim_t = Left(Right(raw_hora_fim,8),5)

										prec_lote = FormatNumber(RS_Lotes("Valor"),2)

									%>

									<!-- <tr <%If iNow mod 2 = 0 Then %> style="background:#fff;" <%End If%>>
										<td><strong><%=RS_Lotes("Nome")%></strong></td>
										<td><strong>R$ <%=prec_lote%></strong></td>
										<td><strong><%=hora_fim_n%> às <%=hora_fim_t%> </strong></td>
									</tr> -->

									<%
										RS_Lotes.MoveNext
										iNow = iNow + 1
										Wend
										RS_Lotes.Close
									End If
									%>



								</table>


                                <div style="float: left; width: 100%">
									<a href="#finalizar_pedido" onclick="ConfirmarCompra()"><div class="bt_fechar_pedido trn" style="float: right" data-trn-key="text17">Concluir a compra</div></a>
								</div>
                                <%End If%>

								<div style="float: left; width: 100%; margin-top: 10px; padding: 10px 0 0; font-weight: 100; border-top: 1px dotted #999" class="trn" data-trn-key="text18">
								Se quiser comprar ingressos para outras pessoas, utilize o quadro de busca abaixo. A busca deverá ser feita pelo número do <font style="font-weight: bold">CPF</font> ou <font style="font-weight: bold">E-mail</font>, em caso de estrangeiros.<br>
								<font style="font-size: 10px"><em>Obs.: para que o <strong>CPF</strong> ou <strong>E-mail</strong> constem em nossa base de dados, é necessário que estas pessoas já tenham feito seu credenciamento.</em></font>
								</div>

								<label style="width:400px; margin-left: -5px;">
									<div style="width: 400px;" class="trn" data-trn-key="text19">BUSCA:</div>
									<input id="formBusca" type="text" maxlength="100" max="100" style="width:200px; padding: 1px; height: 18px;" name="frmID_Visitante">
									<a href="#buscar" onclick="buscar_visitante()"><div class="bt_buscar" style="float: left;">Concluir a compra</div></a>
								</label>

								<script>
									function buscar_visitante(){
										show_loading();
										var timeout = setTimeout(
											function (){
												alert('Tempo de resposta de 15 seg. excedido.\n\nFavor tentar novamente ou reiniciar seu processo.\n\nti@btsmedia.biz');
											}
										, 15000);
										$("#DadosVisitante").html('');
										$("#IdVisitante").val('');
										$("#NascVisitante").val('');
										$("#NascVisitante").val('');
										$("#ResultadoBusca").html('');
										$("#TelaResultado").hide();

										if ($("#formBusca").val()=='') {
											$("#loading").fadeOut();
											clearTimeout(timeout);
											Erros_Busca(3);
											$("#formBusca").val('');
											$("#formBusca").focus();

										} else {
											//jAlert("/tickets/busca.asp?busca=" + $("#formBusca").val() + "&pedido=" + <%=ID_Pedido%>,"URL");
											$.ajax({
												url: "/tickets/busca.asp?busca=" + $("#formBusca").val() + "&pedido=" + <%=ID_Pedido%>,
												success: function(data){

													Resposta = data.split(';');
													//alert(Resposta[0])

													if (Resposta[0]=='Erro') {
														$("#loading").fadeOut();
														clearTimeout(timeout);
														Erros_Busca(Resposta[1]);
														$("#formBusca").val('');
														$("#formBusca").focus();
													} else {
														$("#loading").fadeOut();
														clearTimeout(timeout);
														$("#DadosVisitante").html("<div style='width: 570px; float: left; font-size: 12px; border-top: 1px dotted #999; padding: 5px;'><div style='float: left; width: 470px; padding: 5px;'>" + Resposta[2] + "</div><div style='float: right; width: 30px; padding: 5px;'><a href='#add' onclick='ValidarData();'><img src='/img/forms/add.png' alt='Adicionar este Visitante ao minha Compra!' title='Adicionar este Visitante ao minha Compra!'></a></div></div>");
														$("#IdVisitante").val(Resposta[0]);
														$("#IRC").val(Resposta[1]);
														$("#NascVisitante").val(Resposta[3]);
														$("#ResultadoBusca").html('Resultado da Busca:<div style="float: right; width: 70px; padding: 5px;">Adicionar</div>');
														$("#TelaResultado").show();
														$("#formBusca").val('');
													}
												}
											});
										}
									}

									function Erros_Busca(valor){
										var langMsg = readCookie("lang");
										if(valor==0){
												if(langMsg == null)
													jAlert('Este <strong>CPF</strong> não está cadastrado em nosso banco de dados. Para efetuar a compra de ingressos, a pessoa dona deste <strong>CPF</strong> deverá efetuar seu credenciamento previamente.','CPF não encontrado!');
												else if(langMsg == "es")
													jAlert('Este <strong>Número de identificación fiscal (CPF)</strong> no está registrado en nuestro banco de datos. Para realizar compra de billetes, el titular de este <strong>Número de identificación fiscal (CPF)</strong> deberá realizar su acreditación previamente.','Número de identificación fiscal (CPF) no encontrado!');
												else if(langMsg == "en")
													jAlert('This <strong>CPF</strong> is not registered in our database. To make the ticket purchase, the person, who owns this  <strong>CPF</strong> , shall register himself/herself in advance.','CPF not found!');
												else
													jAlert('Este <strong>CPF</strong> não está cadastrado em nosso banco de dados. Para efetuar a compra de ingressos, a pessoa dona deste <strong>CPF</strong> deverá efetuar seu credenciamento previamente.','CPF não encontrado!');
												
										}else if(valor==1){
												if(langMsg == null)
													jAlert('Este <strong>CPF</strong> já comprou ingresso em outro Pedido!','CPF encontrado!');
												else if(langMsg == "es")
													jAlert('¡Este <strong>número de identificación fiscal (CPF)</strong> ya ha comprado el billete en otro Pedido!','Número de identificación fiscal (CPF) encontrado!');
												else if(langMsg == "en")
														jAlert('This <strong>CPF</strong> has already purchased a ticket in another Purchase order!','CPF found!');
												else	
													jAlert('Este <strong>CPF</strong> já comprou ingresso em outro Pedido!','CPF encontrado!');
										}else if(valor==2){
												if(langMsg == null)
													jAlert('Este <strong>CPF</strong> já está em seu Pedido!','CPF encontrado!');
												else if(langMsg == "es")
													jAlert('¡Este <strong>número de identificación fiscal (CPF)</strong> ya está en su Pedido!','Número de identificación fiscal (CPF) encontrado!');
												else if(langMsg == "en")
													jAlert('This <strong>CPF</strong> is already in your Purchase order!','CPF found!');
												else
													jAlert('Este <strong>CPF</strong> já está em seu Pedido!','CPF encontrado!');
										}else if(valor==3){
											jAlert('Campo <strong>OBRIGATÓRIO</strong>.<br>Digite um CPF ou Passaporte para localizar um visitante!','Aviso!');
										}else if(valor==4){
											jAlert('Você já comprou seu ingresso!','CPF encontrado!');
										}else if(valor==5){
											jAlert('Este CPF ou e-mail já validou uma cortesia!','CPF/E-mail encontrado!');
										}
									}

									function ValidarData(){
										var langMsg = readCookie("lang");
										if(langMsg == null)
										{
												jPrompt('Por motivo de segurança, digite a <strong>DATA DE NASCIMENTO (dd/mm/aaaa)</strong> da pessoa que você quer adicionar ao seu <strong>Pedido</strong>','','Confirmação de Dados!', function(data){
												MontaData = data.split("/");
												//alert(data);

												NascimentoA = MontaData[0] + MontaData[1] +  MontaData[2];
												//alert(NascimentoA);

												NascimentoB = $("#NascVisitante").val();
												//alert(NascimentoB);
												if (NascimentoA == NascimentoB ) {
													$("#form_pedidos").submit();
												} else {
													jAlert('Os dados não conferem.<br><br>Tente novamente.','Erro!');
												}
											});
										}
										else if(langMsg == "es"){
												jPrompt('Por motivo de seguridad, ingrese al <strong>FECHA DE NACIMIENTO (dd/mm/aaaa)</strong> de la persona que desea agregar a su <strong>Pedido</strong>','','Confirmar los datos!', function(data){
												MontaData = data.split("/");
												//alert(data);

												NascimentoA = MontaData[0] + MontaData[1] +  MontaData[2];
												//alert(NascimentoA);

												NascimentoB = $("#NascVisitante").val();
												//alert(NascimentoB);
												if (NascimentoA == NascimentoB ) {
													$("#form_pedidos").submit();
												} else {
													jAlert('Los datos no coinciden.<br><br>Intente de nuevo.','Erro!');
												}
											});
										}
										else if(langMsg == "en"){
												jPrompt('For security reasons, type the <strong>DATE OF BIRTH (day/month/year)</strong> of the person, who you want to add to your <strong>Purchase order</strong>','','Data confirmation!', function(data){
												MontaData = data.split("/");
												//alert(data);

												NascimentoA = MontaData[0] + MontaData[1] +  MontaData[2];
												//alert(NascimentoA);

												NascimentoB = $("#NascVisitante").val();
												//alert(NascimentoB);
												if (NascimentoA == NascimentoB ) {
													$("#form_pedidos").submit();
												} else {
													jAlert('Data does not match.<br><br>Try again.','Error!');
												}
											});
										}
										else
										{
											jPrompt('Por motivo de segurança, digite a <strong>DATA DE NASCIMENTO (dd/mm/aaaa)</strong> da pessoa que você quer adicionar ao seu <strong>Pedido</strong>','','Confirmação de Dados!', function(data){
												MontaData = data.split("/");
												//alert(data);

												NascimentoA = MontaData[0] + MontaData[1] +  MontaData[2];
												//alert(NascimentoA);

												NascimentoB = $("#NascVisitante").val();
												//alert(NascimentoB);
												if (NascimentoA == NascimentoB ) {
													$("#form_pedidos").submit();
												} else {
													jAlert('Os dados não conferem.<br><br>Tente novamente.','Erro!');
												}
											});
										}
										}

										function ConfirmarCompra(){
												var langMsg = readCookie("lang");
												if(langMsg == null)
													jConfirm('Você deseja finalizar o seu Pedido e realizar o Pagamento?','Finalizar Pedido?', function(data){if(data==true){$("#FinalizarPedido").submit();}});
												else if(langMsg == "es")
													jConfirm('¿Usted quiere concluir su Pedido y realizar el Pago?','Concluir pedido?', function(data){if(data==true){$("#FinalizarPedido").submit();}});
												else if(langMsg == "en")
													jConfirm('Do you want to finalize your Purchase order and make the payment?','Finalize purchase order?', function(data){if(data==true){$("#FinalizarPedido").submit();}});
												else
													jConfirm('Você deseja finalizar o seu Pedido e realizar o Pagamento?','Finalizar Pedido?', function(data){if(data==true){$("#FinalizarPedido").submit();}});
										}

										function RemoverVisitante(ID,PEDIDO){
											var langMsg = readCookie("lang");
											if(langMsg == null)
												jConfirm('Você tem certeza que deseja remover esta pessoa?','Remover Visitante?', function(data){if(data==true){window.location = '/tickets/pedido.asp?acao=remover&aceito=1&id='+ID+'&pedido='+PEDIDO;}});
											else if(langMsg == "es")
												jConfirm('¿Está seguro que desea eliminar esta persona?','¿Eliminar visitante?', function(data){if(data==true){$("#FinalizarPedido").submit();}});
											else if(langMsg == "en")
												jConfirm('Are you sure that you want to remove this person?','Remove visitor?', function(data){if(data==true){$("#FinalizarPedido").submit();}});
											else
												jConfirm('Você tem certeza que deseja remover esta pessoa?','Remover Visitante?', function(data){if(data==true){window.location = '/tickets/pedido.asp?acao=remover&aceito=1&id='+ID+'&pedido='+PEDIDO;}});
										}
								</script>

								<div id="TelaResultado" style="display: none; float: left; width: 580px;">
									<div id="ResultadoBusca" style="margin-top: 5px; width: 575px; float: left; border-top: 1px dotted #999; font-size: 14px; padding: 5px 0 5px 5px; background: #dadada"></div>
									<div id="DadosVisitante"></div>
									<input type="hidden" value="" id="IdVisitante" name="IdVisitante">
									<input type="hidden" value="" id="IRC" name="IRC">
									<input type="hidden" value="" id="NascVisitante" name="NascVisitante">
									<input type="hidden" value="<%=ID_Pedido%>" id="ID_Pedido" name="ID_Pedido">
									<input type="hidden" value="adicionar" id="acao" name="acao">
								</div>
                        <%End If%>
                        </div>
                    </fieldset>
                </form>

                <form action="/tickets/pagamento.asp" method="post" name="FinalizarPedido" id="FinalizarPedido">
                	<input type="hidden" value="<%=ID_Pedido%>" id="IDPedido" name="IDPedido">
                </form>
            <br/>

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
<%
Conexao.Close
%>
