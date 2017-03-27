<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/admin/inc/gravar_limpar_texto.asp"-->
<%
Response.Expires = -1
Response.AddHeader "Cache-Control", "no-cache"
Response.AddHeader "Pragma", "no-cache"
response.Charset = "utf-8"
response.ContentType = "text/html"

'=======================================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
			  Conexao.Open Application("cnn")
'=======================================================================


if ((Not Session("pedido") = Limpar_Texto(Request("pedido"))) And Limpar_Texto(Request("pedido")) <> "") Or Session("pedido") = Limpar_Texto(Request("pedido")) Then
	Session("pedido") = Limpar_Texto(Request("pedido"))
	Response.Redirect("recuperar_confirmacao_pagamento.asp")
Else
	Session("aux") = "1"
	'Response.Write(Session("pedido"))
End If

if Session("aux") <> "1" Then
	Valor = Session("pedido")
	If Session("pedido") <> "" Then
	    Response.Redirect("recuperar_confirmacao_pagamento.asp")
	Else
		response.Redirect("http://www.mbxeventos.net/AOLABF2017/")
	End If
End If

	Idioma 	= Session("cliente_idioma")
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

' Buscar imagens da Feira
	SQL_Edicoes_Configuracao = 	"SELECT " &_
								"	EC.Logo_Email, " &_
								"	EC.Logo_Box, " &_
								"	E.Nome_" & SgIdioma & " as Feira, " &_
								"	EE.Ano as Ano " &_
								"FROM " &_
								"	Edicoes_Configuracao as EC " &_
								"INNER JOIN " &_
								"	Eventos_Edicoes as EE " &_
								"	ON EC.ID_Edicao = EE.ID_Edicao " &_
								"INNER JOIN " &_
								"	Eventos as E" &_
								"	ON EE.ID_Evento = E.ID_Evento " &_
								"WHERE " &_
								"	EC.ID_Edicao = " & Session("cliente_edicao") & " " &_
								"	AND EC.Ativo = 1"
	'response.write(SQL_Edicoes_Configuracao)
	'response.end
	Set RS_Edicoes_Configuracao = Server.CreateObject("ADODB.Recordset")
	RS_Edicoes_Configuracao.CursorType = 0
	RS_Edicoes_Configuracao.LockType = 1
	RS_Edicoes_Configuracao.Open SQL_Edicoes_Configuracao, Conexao

	Logo_email	= RS_Edicoes_Configuracao("Logo_Email")
	Logo_box	= RS_Edicoes_Configuracao("Logo_Box")
	Feira		= RS_Edicoes_Configuracao("Feira")
	Ano			= RS_Edicoes_Configuracao("Ano")
	RS_Edicoes_Configuracao.Close


	SQL_Consulta_Pedidos =	"Select " &_
							"	P.*, " &_
							"	PH.* " &_
							"From Pedidos As P " &_
							"Inner Join Pedidos_Historico as PH " &_
							"	On P.Numero_Pedido = PH.Numero_Pedido " &_
							"Where " &_
							"	P.Numero_Pedido = '" & Session("pedido") & "'  " &_
							"	And P.Status_Pedido = 3" &_
							"	And PH.Status_Pagamento = 1"
	'Response.Write(SQL_Consulta_Pedidos)
	'Response.End
	Set RS_Consulta_Pedidos = Server.CreateObject("ADODB.Recordset")
	RS_Consulta_Pedidos.Open SQL_Consulta_Pedidos, Conexao, 3, 3

	If Not RS_Consulta_Pedidos.Eof Then

		Tickets 			= True
		Numero_Pedido 		= RS_Consulta_Pedidos("Numero_Pedido")
		Numero_Transacao 	= RS_Consulta_Pedidos("Numero_Transacao")
		Codigo_Paypal 		= RS_Consulta_Pedidos("Codigo_Paypal")
		Cod_Autorizacao		= RS_Consulta_Pedidos("Codigo_Autorizacao")
		Valor_Pedido		= FormatNumber(RS_Consulta_Pedidos("Valor_Pedido"),2)
		ID_Visitante		= RS_Consulta_Pedidos("ID_Visitante")
		ID_Pedido 			= RS_Consulta_Pedidos("ID_Pedido")
	End If

	'ID_Pedido = "1"

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>Credenciamento <%=Year(Now())%> - BTS Informa</title>
	<link href="/css/confirmacao_pagamento.css" rel="stylesheet" type="text/css"/>
	<script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
	<script language="javascript" src="/js/funcoes_gerais.js"></script>
	<script language="javascript" src="/js/jquery.translate.js"></script>
	<script language="javascript">
		$(function() {

		  var t = {
			text1: {
			  pt: "Confirmação de Compra do Ingresso",
			  en: "Ticket purchase confirmation",
			  es: "Confirmación de compra del billete"
			},
			text2: {
			  pt: "Pagamento:",
			  en: "Payment",
			  es: "Pago"
			},
			text3: {
			  pt: "Aprovado",
			  en: "Approved",
			  es: "Aprobado"
			},
			text4: {
			  pt: "Numero do Pedido:",
			  en: "Purchase order number:",
			  es: "Numero del Pedido:"
			},
			text5: {
			  pt: "Codigo PayPal:",
			  en: "PayPal Code:",
			  es: "Codigo PaypPal:"
			},
			text6: {
			  pt: "Transa&ccedil;&atilde;o:",
			  en: "Transaction:",
			  es: "Transacci&oacute;n:"
			},
			text7: {
			  pt: "C&oacute;d. da Autoriza&ccedil;&atilde;o:",
			  en: "Authorization Code:",
			  es: "C&oacute;d. de la Autorizaci&oacute;n:"
			},
			text8: {
			  pt: "Valor Pago:",
			  en: "Amount paid:",
			  es: "Cantidad Pagada:"
			},
			text9: {
			  pt: "NOME COMPLETO:",
			  en: "FULL NAME:",
			  es: "NOMBRE COMPLETO:"
			},
			text10: {
			  pt: "TIPO",
			  en: "TYPE",
			  es: "TIPO"
			},
			text11: {
			  pt: "DOCUMENTO",
			  en: "DOCUMENT",
			  es: "DOCUMENTO"
			},
			text12: {
			  pt: "- Para retirar seu ingresso e a credencial para acesso ao evento, tenha em mãos seu comprovante de compra e seu CPF.<br>- O ingresso é pessoal e intransferível, sendo obrigatória a apresentação do CPF para sua retirada.<br>- Não será permitida a entrada de pessoas trajando bermudas, camiseta regata e/ou chinelos.<br>- Proibida a entrada de menores de 16 anos desacompanhados.<br><br/>",
			  en: "- To take your ticket and credential to access the event, you must have your purchase receipt and your CPF.<br>- The ticket is personal and non-transferable, and presentation of the CPF for its withdraw is obligatory.<br>- People are not allowed to enter wearing shorts, tank tops and/or slippers.<br>- Unaccompanied minors under the age of 16 are prohibited to enter<br><br/>",
			  es: "- Para retirar su billete y credencial de identificación para el evento, tenga a mano su comprobante de compra y su número de identificación fiscal (CPF).<br>- El billete es personal e intransferible, es obligatoria la presentación del número de identificación fiscal (CPF) para su retirada.<br>- No se permitirá el ingreso de personas que visten pantalones cortos, camiseta sin mangas y/o chancletas.<br>- Está prohibido la entrada de menores de 16 años sin acompañante<br><br/>"
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
</head>
<body style="margin:10px;">
<table width='640' border='0' cellpadding='0' cellspacing='0'>
    <tr>
        <td><img src='http://credenciamento.btsinforma.com.br/img/geral/informa_exhibition.png' alt=''  hspace='15' /></td>
		<!--td><img src='http://ws.homologabts.com.br/images/informa_exhibition.png' alt='' width='95' height='52' hspace='15' /></td-->
        <td width='15'>&nbsp;</td>
        <td align='right'><img src='http://credenciamento.btsinforma.com.br<%=logo_box%>' alt="<%=Feira%>&nbsp;<%=Ano%>" title="<%=Feira%>&nbsp;<%=Ano%>"/></td>
		<!--td align='right'><img src='http://ws.homologabts.com.br<%=logo_box%>' alt="<%=Feira%>&nbsp;<%=Ano%>" title="<%=Feira%>&nbsp;<%=Ano%>"/></td-->
    </tr>
</table>
<div style="width:600px; text-align:center;">
	<h1><font class="trn" data-trn-key="text1">Confirmação de Compra do Ingresso</font> <br /><%=Feira%>&nbsp;<%=Ano%></h1>
</div>
	<div style="padding: 5px 0; width: 180px; float: left" class="trn" data-trn-key="text2">Pagamento:</div>									<div style="padding: 5px 0; font-weight: 900" class="trn" data-trn-key="text3">Aprovado</div>
	<div style="padding: 5px 0; width: 180px; float: left" class="trn" data-trn-key="text4">Numero do Pedido:</div>							<div style="padding: 5px 0; font-weight: 900"><%=Numero_Pedido%></div>
	<div style="padding: 5px 0; width: 180px; float: left" class="trn" data-trn-key="text5">Codigo Paypal:</div>								<div style="padding: 5px 0; font-weight: 900"><%=Codigo_Paypal%></div>
	<div style="padding: 5px 0; width: 180px; float: left" class="trn" data-trn-key="text6">Transa&ccedil;&atilde;o:</div> 					<div style="padding: 5px 0; font-weight: 900"><%=Numero_Transacao%></div>
	<div style="padding: 5px 0; width: 180px; float: left" class="trn" data-trn-key="text7">C&oacute;d. da Autoriza&ccedil;&atilde;o:</div> 	<div style="padding: 5px 0; font-weight: 900"><%=Cod_Autorizacao%></div>
    <%If CStr(Session("cliente_visitante")) <> CStr(ID_Visitante) Then%>
		<div style="padding: 5px 0; width: 180px; float: left" class="trn" data-trn-key="text8">Valor Pago:</div> 								<div style="padding: 5px 0; font-weight: 900">R$ <%=FormatNumber(Application("Valor_Ticket"),2)%></div>
    <%Else%>
    	<div style="padding: 5px 0; width: 180px; float: left" class="trn" data-trn-key="text8">Valor Pago:</div> 								<div style="padding: 5px 0; font-weight: 900">R$ <%=Valor_Pedido%></div>
    <%End If%>
<br/><br/>
<table>
	<tr>
    	<td style=" border-bottom: 1px dotted #ccc"><b class="trn" data-trn-key="text9">NOME COMPLETO</b></td>
        <td style=" border-bottom: 1px dotted #ccc"><b class="trn" data-trn-key="text10">TIPO</b></td>
        <td style=" border-bottom: 1px dotted #ccc"><b class="trn" data-trn-key="text11">DOCUMENTO</b></td>
    </tr>
<%

	'Response.Write(CStr(Session("cliente_visitante")) <> CStr(ID_Visitante))

	If CStr(Session("cliente_visitante")) <> CStr(ID_Visitante) Then

		SQL_Carrinho = 	"Select " &_
						"	C.ID_Carrinho,  " &_
						"	C.ID_Visitante,  " &_
						"	C.ID_Pedido,  " &_
						"	C.ID_Rel_Cadastro, " &_
						"	C.ID_Rel_Cadastro, " &_
						"	P.Status_Pedido, " &_
						"	V.Nome_Completo, " &_
						"	V.CPF, " &_
						"	V.Passaporte " &_
						"From  Pedidos_Carrinho  As C " &_
						"Inner Join Visitantes As V On V.ID_Visitante = C.ID_Visitante " &_
						"Inner Join Pedidos As P On P.ID_Pedido = C.ID_Pedido " &_
						"Where " &_
						"	P.Numero_Pedido = '" & Session("pedido") & "' "&_
						"	And (P.ID_Visitante = '" & Session("cliente_visitante") & "' " &_
						"	Or C.ID_Visitante = '" & Session("cliente_visitante") & "') " &_
						"	And C.Cancelado = 0"
	Else
		SQL_Carrinho = 	"Select " &_
						"	C.ID_Carrinho,  " &_
						"	C.ID_Visitante,  " &_
						"	C.ID_Pedido,  " &_
						"	C.ID_Rel_Cadastro, " &_
						"	P.Status_Pedido, " &_
						"	V.Nome_Completo, " &_
						"	V.CPF, " &_
						"	V.Passaporte " &_
						"From  Pedidos_Carrinho  As C " &_
						"Inner Join Visitantes As V On V.ID_Visitante = C.ID_Visitante " &_
						"Inner Join Pedidos As P On P.ID_Pedido = C.ID_Pedido " &_
						"Where " &_
						"	P.Numero_Pedido = '" & Session("pedido") & "' " &_
						"	And C.Cancelado = 0"

	End if

	'Response.Write(SQL_Carrinho)

	Set RS_Carrinho = Server.CreateObject("ADODB.Recordset")
	RS_Carrinho.Open SQL_Carrinho, Conexao, 3, 3



Primeiro = 0
Z = True

If Not RS_Carrinho.Eof Then
    While Not RS_Carrinho.Eof
%>

	<tr bgcolor="<%=Cor_Fundo%>" style="padding: 5px; font-weight: 100">
    	<td style="padding: 5px; width: 375px; border-bottom: 1px dotted #ccc"><%=RS_Carrinho("Nome_Completo")%></td>
        <td style="padding: 5px; width: 100px; border-bottom: 1px dotted #ccc"><%If Len(Trim(RS_Carrinho("CPF"))) > 0 Then%>CPF<%Else%>Passaporte<%End If%></td>
        <td style="padding: 5px; width: 100px; border-bottom: 1px dotted #ccc"><%If Len(Trim(RS_Carrinho("CPF"))) > 0 Then Response.Write(RS_Carrinho("CPF")) Else Response.Write(RS_Carrinho("Passaporte"))%></td>
    </tr>

<%
    RS_Carrinho.MoveNext
    Wend
End If
%>
</table>
<br/>
<div style="width:600px; text-align:center;">
<!--img src='http://credenciamento.btsinforma.com.br/img/geral/logos/Feira_e-commerce.jpg' title="<%=Feira%>&nbsp;<%=Ano%>"/-->
<img src='http://credenciamento.btsinforma.com.br/img/geral/logos/faixa_abf_2017.png' title="<%=Feira%>&nbsp;<%=Ano%>"/>
</div>
<br />
<div id="footer" class="trn" data-trn-key="text12">
- Para retirar seu ingresso e a credencial para acesso ao evento, tenha em mãos seu comprovante de compra e seu CPF.<br>
- O ingresso é pessoal e intransferível, sendo obrigatória a apresentação do CPF para sua retirada.<br>
- Não será permitida a entrada de pessoas trajando bermudas, camiseta regata e/ou chinelos.<br>
- Proibida a entrada de menores de 16 anos desacompanhados.<br>

<br/>
</div>
</body>

</html>
<%
Conexao.Close
%>
