<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Response.Charset="ISO-8859-1"%>
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/includes/texto_caixaAltaBaixa.asp"-->
<!--#include virtual="/scripts/enviar_email_senha.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=uiso-8859-1" />
<title>Credenciamento <%=Year(Now())%> - BTS Informa</title>

<!-- Script desta página FIM -->
<%

session("teste_paypal") = false
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================

Numero_Transacao = Limpar_Texto(Request("transacao"))
Numero_Pedido	= Limpar_Texto(Request("pedido"))

Function URLDecode(sConvert)
    Dim aSplit
    Dim sOutput
    Dim I
    If IsNull(sConvert) Then
       URLDecode = ""
       Exit Function
    End If

    ' convert all pluses to spaces
    sOutput = REPLACE(sConvert, "+", " ")

    ' next convert %hexdigits to the character
    aSplit = Split(sOutput, "%")

    If IsArray(aSplit) Then
      sOutput = aSplit(0)
      For I = 0 to UBound(aSplit) - 1
        sOutput = sOutput & _
          Chr("&H" & Left(aSplit(i + 1), 2)) &_
          Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
      Next
    End If

    URLDecode = sOutput
End Function

'For Each item In Request.Form
'	Response.Write "" & item & " - Value: " & Request.Form(item) & "<BR />"
'Next

SQL_Consulta_Pedido = 	"Select " &_
						"	P.* " &_
						"	,H.*  " &_
						"From Pedidos As P " &_
						"Inner Join Pedidos_Historico As H " &_
						"	On P.Numero_Pedido = H.Numero_Pedido " &_
						"	And H.Numero_Transacao = '" & Numero_Transacao & "' " &_
						"Where " &_
						"	P.Numero_Pedido = '" & Numero_Pedido & "'"
	'
Set RS_Consulta_Pedido = Server.CreateObject("ADODB.Recordset")
RS_Consulta_Pedido.Open SQL_Consulta_Pedido, Conexao, 3, 3
'response.write SQL_Consulta_Pedido
'response.end
If Not RS_Consulta_Pedido. Eof Then

	Idioma 			= RS_Consulta_Pedido("ID_Idioma")
	ID_Visitante 	= RS_Consulta_Pedido("ID_Visitante")
	ID_Rel_Cadastro = RS_Consulta_Pedido("ID_Rel_Cadastro")
	Retorno_Pedido	= RS_Consulta_Pedido("Retorno")
	Aprovacao = RS_Consulta_Pedido("Status_Pagamento")


	SQL_Cadastro_Visitantes =	"Select " &_
								"	RC.ID_Relacionamento_Cadastro as IRC " &_
								"	,RC.ID_Tipo_Credenciamento " &_
								"	,TC.ID_Formulario " &_
								"	,RC.ID_Edicao " &_
								"	,RC.ID_Empresa " &_
								"	,RC.ID_Visitante " &_
								"	,V.Nome_Completo " &_
								"	,V.CPF " &_
								"	,V.Senha " &_
								"	,RC.ID_Visitante " &_
								"	,V.Email " &_
								"From Relacionamento_Cadastro as RC " &_
								"Inner Join Tipo_Credenciamento as TC ON TC.ID_Tipo_Credenciamento = RC.ID_Tipo_Credenciamento " &_
								"Inner Join Visitantes as V ON V.ID_Visitante = RC.ID_Visitante " &_
								"Where  " &_
								"	TC.ID_Idioma = " & Idioma & "		/* Idioma	*/ " &_
								"	AND V.ID_Visitante = '" & ID_Visitante & "'  " &_
								"	AND RC.ID_Relacionamento_Cadastro = '" & ID_Rel_Cadastro & "'"
	'Response.Write(SQL_Cadastro_Visitantes)
	'response.end
	Set RS_Cadastro_Visitantes = Server.CreateObject("ADODB.Recordset")
	RS_Cadastro_Visitantes.Open SQL_Cadastro_Visitantes, Conexao, 3, 3

	ID_Edicao               = RS_Cadastro_Visitantes("ID_Edicao")
	ID_TP_Credenciamento    = RS_Cadastro_Visitantes("ID_Tipo_Credenciamento")
	TP_Formulario           = RS_Cadastro_Visitantes("ID_Formulario")
	IRC                     = RS_Cadastro_Visitantes("IRC")
	ID_Empresa              = RS_Cadastro_Visitantes("ID_Empresa")
	ID_Visitante            = RS_Cadastro_Visitantes("ID_Visitante")
	Nome_Visitante          = RS_Cadastro_Visitantes("Nome_Completo")
	CPF_Visitante           = RS_Cadastro_Visitantes("CPF")
	Email_Visitante           = RS_Cadastro_Visitantes("Email")

	Session("cliente_edicao")		= ID_Edicao
	Session("cliente_idioma")		= Idioma
	Session("cliente_tipo")			= ID_TP_Credenciamento
	Session("cliente_formulario")	= TP_Formulario
	Session("cliente_logado")		= 1
	Session("cliente_visitante")	= ID_Visitante

    Session("cliente_cadastro")     = IRC
    Session("cliente_empresa")      = ID_Empresa
    Session("cliente_nome")         = Nome_Visitante
    Session("cliente_cpf")			= CPF_Visitante

	If Session("cliente_edicao") = "" OR Session("cliente_idioma") = "" OR Session("cliente_logado") = "" or Session("cliente_visitante") = "" Then
	  response.Redirect("http://www.mbxeventos.net/AOLABF2017/")
	End If


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
						"	ID_Idioma = " & Idioma & " " &_
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
	<% If Request("teste") = "s" Then %>
		<!--#include virtual="/includes/exibir_array.asp"-->
	<% End IF

		' Select IMG Faixa
		SQL_Img_Faixa 	=	"Select " &_
							"	Img_Faixa " &_
							"From Tipo_Credenciamento " &_
							"Where ID_Tipo_Credenciamento = " & ID_TP_Credenciamento
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

		' Tratar variáveis de RETORNO
		'====================================



		'===================================
		SQL_Consulta_Pedidos = 	"Select " &_
								"	P.* " &_
								"From " &_
								"	Pedidos As P " &_
								"Where " &_
								"	P.ID_Edicao = '" & Session("cliente_edicao") & "' " &_
								"	And P.ID_Rel_Cadastro = '" & Session("cliente_cadastro") & "' " &_
								"	And P.ID_Visitante = '" & Session("cliente_visitante")  & "' "' &_
'										"	And P.Status_Pedido = 1"
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


		If Aprovacao = "True" Then
			'=============================================
			' Atualizar Pedido como FINALIZADO
			SQL_Atualizar_Pedido = 	"Update Pedidos " &_
									"Set Status_Pedido = 3  " &_
									"Where Numero_Pedido = '" & session("Numero_Pedido") & "'"



			Conexao.Execute(SQL_Atualizar_Pedido)
			'=============================================
			' Realizar LOOP nos Itens e marcar como Cancelado caso Exista em outro Carrinho não finalizado
			SQL_Listar_Carrinho = 	"Select " &_
									"	ID_Visitante " &_
									"	,ID_Pedido " &_
									"From Pedidos_Carrinho " &_
									"Where " &_
									"	ID_Pedido = " & ID_Pedido
		'response.write("<hr>" & SQL_Listar_Carrinho & "<hr>")
		'Response.Write(SQL_Atualizar_Pedido)

			Set RS_Listar_Carrinho = Server.CreateObject("ADODB.Recordset")
			RS_Listar_Carrinho.Open SQL_Listar_Carrinho, Conexao, 3, 3

			If not RS_Listar_Carrinho.BOF or not RS_Listar_Carrinho.EOF Then
				' Loop dos visitantes que foram comprados
				While not RS_Listar_Carrinho.EOF
					' Dados
					ID_Visitante = RS_Listar_Carrinho("ID_Visitante")
					ID_Pedido = RS_Listar_Carrinho("ID_Pedido")

					'============================================
					'Selecionar Visitante em outros Carrinhos
					SQL_Visitantes_Outro_Carrinho =	"Select " &_
													"	C.ID_Pedido " &_
													"	,C.ID_Visitante " &_
													"	,P.Status_Pedido " &_
													"From Pedidos_Carrinho as C " &_
													"Inner Join Pedidos as P ON P.ID_Pedido = C.ID_Pedido " &_
													"Where	C.ID_Visitante 	 = " & ID_Visitante & " /* Visitante à localizar */ " &_
													"	AND	C.ID_Pedido 	<> " & ID_Pedido & " /* Meu Pedido Atual */ " &_
													"	AND P.Status_Pedido = 1 /* Pedido não finalizado */ "
					Set RS_Visitantes_Outro_Carrinho = Server.CreateObject("ADODB.Recordset")
					RS_Visitantes_Outro_Carrinho.Open SQL_Visitantes_Outro_Carrinho, Conexao, 3, 3

					'============================================
					' Se encontrar em outro carrinho
					If not RS_Visitantes_Outro_Carrinho.BOF or not RS_Visitantes_Outro_Carrinho.EOF Then

						While not RS_Visitantes_Outro_Carrinho.EOF
							ID_Pedido_Atualizar = RS_Visitantes_Outro_Carrinho("ID_Pedido")

							'============================================
							' Marcar Visitante como cancelado se estiver em outro pedido NÃO FINALIZADO
							SQL_Atualizar_para_Cancelado = 	"Update Pedidos_carrinho " &_
															"Set Cancelado = 1 " &_
															"Where	" &_
															"		ID_Visitante = " & ID_Visitante & " " &_
															"	AND	ID_Pedido = " & ID_Pedido_Atualizar

							Conexao.Execute(SQL_Atualizar_para_Cancelado)

							'============================================
							' Diminuir o valor do pedido em 1 Item
							Valor_Ticket = Application("Valor_Ticket")

							SQL_Atualizar_Pedido = 	"Update Pedidos " &_
													"Set	Valor_Pedido = Valor_Pedido - " & Valor_Ticket & " " &_
													"Where ID_Pedido = " & ID_Pedido_Atualizar

							Conexao.Execute(SQL_Atualizar_Pedido)

							RS_Visitantes_Outro_Carrinho.MoveNext
						Wend
						RS_Visitantes_Outro_Carrinho.Close
					End If

					'============================================

					RS_Listar_Carrinho.MoveNext
				Wend
				RS_Listar_Carrinho.Close
			End If
		End If




	%>
	<link href="/css/base_forms.css" rel="stylesheet" type="text/css" />
	<link href="/css/estilos.css" rel="stylesheet" type="text/css">
	<link href="/css/jquery.alerts.css" rel="stylesheet" type="text/css">

	<script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
	<script language="javascript" src="/js/jquery-ui-1.8.7.core_eff-slide.js"></script>
	<script language="javascript" src="/js/jquery.alerts.js"></script>
	<script language="javascript" src="/js/validar_forms.js"></script>
	<script language="javascript" src="/js/funcoes_gerais.js"></script>

	<!-- Script desta página -->
	<script language="javascript" src="default.js" charset="utf-8"></script>

	<script language="javascript">
	var idioma_atual = '<%=Session("cliente_idioma")%>';
	var select       = '<%=textos_array(36)(2)%>';
	var cor_fundo 	 = '<%=faixa_cor%>';
	var tp_formulario = '';

	$(document).ready(function(){
		var erro = '<%=Request("erro")%>';
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

	function link(qual, direcao){
		if (direcao == 'voltar') {
			$('#conteudo').css( {"z-index": 10 }).hide("slide", { direction: "right" }, 1000);
		} else if (direcao == 'ir') {
			$('#conteudo').css( {"z-index": 10 }).show("slide", { direction: "right" }, 1000);
		} else {
			$('#conteudo').css( {"z-index": 10 }).hide("slide", { direction: "right" }, 1000);
		}
		setTimeout(function() {
			var urls = '/tickets/' + qual;
			document.location = urls;
		},1000);
	}
	</script>
	</head>

	<body >
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
						<td width="189" height="45" background="/img/geral/faixa_fundo_esqs.gif"><img id="img_faixa_esq" src="/img/geral/tipos/Faixa_Tickets.gif" width="189" height="45"></td>
						<td id="img_fundo_selecionado" height="45" background="<%=faixa_fundo%>" class="atencao_13px cor_branco">
							<div id="txt_1" style="padding-left:20px; float:left; line-height:40px;" align="left"></div>
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
				<!-- Alert error -->
				<div id="aviso_topo" class="fs_12px arial cor_cinza2">
					<img src="/img/forms/alert-icon.png" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso">
					&nbsp;<span id="txt_topo"><!--Por favor preencher corretamente os itens em destaque !--><%=textos_array(39)(2)%></span>
				</div><br/>
				<!-- End Alert error -->

				<table width="850" border="0" cellpadding="0" cellspacing="0">
					<tr>
						<td width="800" height="30" bgcolor="#414042" class="arial fs_13px cor_branco" style="padding-left:15px;"><b>Ol&aacute;</b> <%=Nome_Visitante%></td>
						<td width="500" height="30" bgcolor="#414042" align="right"><img src="/img/botoes/voltar.gif" width="47" height="15" hspace="5" class="cursor" onClick="link('status.asp');"></td>
						<td width="50" height="30" bgcolor="#414042" style=" border-left:#ccc 1px solid;" align="right"><img src="/img/botoes/sair.gif" width="47" height="15" class="cursor" onClick="sair();"></td>
					</tr>
				</table>
				<br/>
				<!--#Include virtual="/tickets/menu_lateral.asp"-->
					<form id="form_pedidos" name="form_pedidos" onsubmit="return false" action="/tickets/pedido.asp" method="post">
						<fieldset style="float: right; width: 580px;">

							<legend>Detalhes do Pagamento:
							</legend><div id="parcAssis" class="div_parceria" style="width:580px; float: right; margin-top: 10px;">


									<div style="font-weight: 100; width: 575px; float: left; font-size: 14px; font-size: 14px; padding: 5px 0 5px 5px;">
										<%
										'Response.Write(Retorno_Pedido & "<br>------------------------------------------<br>")

										If Aprovacao = "True"  Then
										%>


                                            <div style="float:right; padding-right:5px;">

                                            </div>
											<br><br>
                                            <div style="padding: 5px 0; width: 180px; float: left">Pagamento:</div> 								<div style="padding: 5px 0; font-weight: 900">Aprovado</div>
                                            <div style="padding: 5px 0; width: 180px; float: left">Numero do Pedido:</div>							<div style="padding: 5px 0; font-weight: 900"><%=session("Numero_Pedido")%></div>
                                            <div style="padding: 5px 0; width: 180px; float: left">Transa&ccedil;&atilde;o:</div> 					<div style="padding: 5px 0; font-weight: 900"><%="Aprovada"%></div>
                                            <div style="padding: 5px 0; width: 180px; float: left">Token Paypal:</div> 	<div style="padding: 5px 0; font-weight: 900"><%=session("token")%></div>
                                            <div style="padding: 5px 0; width: 180px; float: left">Valor Pago:</div> 								<div style="padding: 5px 0; font-weight: 900"><%If Cint(Idioma) = 1 then Response.Write("R$") Else Response.Write("$")%>&nbsp;<%=FormatNumber(session("finalPaymentAmount"),2)%></div>

										<%
										'Session("cliente_enviar_email") = 0
										'response.Write ID_Visitante
										'response.End
										session("Numero_Pedido") = Request("pedido")
										'If Session("cliente_enviar_email") = "" Or Session("cliente_enviar_email") = 0 Then
											Enviar_Email_Senha Session("cliente_edicao"), Session("cliente_idioma"), "", "", Email_Visitante, "", "", "Enviar_Ticket", session("Numero_Pedido"), ID_Visitante
											'Response.Write Email_Visitante
											'Response.End
											Session("cliente_enviar_email") = True
										'End If

										' Se Nao Aprovado
										Else
										%>
                                        	<br /><br />
											<div style="width: 100px; float: left">Pagamento:</div> <div style="font-weight: 900;">Recusado</div><br>
                                            <div style="width: 100%; float: left"></div><br>
                                            <div style="width: 100%; float: left"><em><strong>Verifique a forma de pagamento na sua conta paypal.</strong></em></div><br><br>
                                        <%End If%>
									</div>
                                    <%
									' Se pedido APROVADO
									If Aprovacao = "True" Then
										%>
										<br /><div style="font-size:16px; color:#1F497D; font-weight:bold; font-family:'Calibri','sans-serif'; text-align:center;">Sua compra foi realizada com sucesso! Retire seu ingresso nos guichês de atendimento na entrada da ABF Franchising Expo 2016 - Expo Center Norte</div><br /><br />

										<div style="float: left; width: 560px; background-color:#fff; font-weight:normal; padding:10px; border-top: 1px dotted #ccc; border-bottom: 1px dotted #ccc; line-height:18px;">
											Caso queira comprar ingresso para outra pessoa, basta clicar em <font style="font-weight: bold"><em>"Novo pedido"</em></font> no menu lateral e realizar uma busca por <font style="font-weight: bold"><em>CPF</em></font> ou <font style="font-weight: bold"><em>passaporte</em></font>,
                                            em caso de estrangeiros.<br>Para que o <font style="font-weight: bold"><em>CPF</em></font> ou <font style="font-weight: bold"><em>Passaporte</em></font> constem em nossa base de dados, &eacute; necess&aacute;rio que estas pessoas j&aacute; tenham feito seu credenciamento.
										</div>
										<%
									End If
									%>
									<br/><br/>
									<div style="float: left; width: 100%">
                                    <%
									'Botão para voltar a compra
									If Aprovacao = "True"  then

										%>
											<!-- Não exibir botão continuar
                                            <a href="/tickets/status.asp"><div id="loader_2" class="bt_meus_pedidos" style="float: right;">Concluir Este Pedido</div></a>
                                            -->
											<a href="#_" onclick="link('novo_pedido.asp','voltar');"><div class="voltar_box" style="float: left">Voltar</div></a>
										<%

									Else
										%>

											<a href="#_" onclick="link('status.asp');"><div class="continuar" style="float: right">Continuar</div></a>
										<%
									End If
									%>
									</div>
							</div>
						</fieldset>
					</form>

					<form action="/tickets/pagamento.asp" method="post" name="FinalizarPedido" id="FinalizarPedido">
						<input type="hidden" value="<%=ID_Pedido%>" id="IDPedido" name="IDPedido">
					</form>
				<br/>
				<%

				%>
				<!-- Alert error -->
				<div id="aviso" class="fs_12px arial cor_cinza2" style="display:inline-table; margin-top:15px;">
					<img src="/img/forms/alert-icon.png" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso">

				</div>
				<!-- End Alert error -->
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

End If
'Conexao.Close
%>
