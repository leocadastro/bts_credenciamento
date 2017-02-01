<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Response.Charset="ISO-8859-1"%>
<!--#include virtual="/includes/limpar_texto.asp"-->
<!--#include virtual="/scripts/ConsultarWebService.asp"-->
<!-- #include file ="paypalfunctions.asp" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content-type:"text/html; charset=ISO-8859-1" />
<title>Credenciamento <%=Year(Now())%> - BTS Informa</title>
<link href="http://credenciamento.btsinforma.com.br/css/base_forms.css" rel="stylesheet" type="text/css" />
<link href="http://credenciamento.btsinforma.com.br/css/estilos.css" rel="stylesheet" type="text/css">

<%
currencyCodeType = "BRL"
paymentType = "Sale"

'------------------------------------
' The returnURL is the location where buyers return to when a
' payment has been succesfully authorized.
'
' This is set to the value entered on the Integration Assistant
'------------------------------------
returnURL = "HTTP://credenciamento.btsinforma.com.br/tickets/retorno.asp"
'returnURL = "HTTP://ws.homologabts.com.br/tickets/retorno.asp"
'returnURL = "http://localhost:81/tickets/retorno.asp"

'------------------------------------
' The cancelURL is the location buyers are sent to when they click the
' return to XXXX site where XXX is the merhcant store name
' during payment review on PayPal
'
' This is set to the value entered on the Integration Assistant
'------------------------------------
cancelURL = "HTTP://credenciamento.btsinforma.com.br/tickets/cancelar.asp"
'cancelURL = "HTTP://ws.homologabts.com.br/tickets/cancelar.asp"
'cancelURL = "http://localhost:81/tickets/pagamento.asp"

'------------------------------------
' Calls the SetExpressCheckout API call
'
' The CallShortcutExpressCheckout function is defined in the file PayPalFunctions.asp,
' it is included at the top of this file.
'-------------------------------------------------
Set resArray = GetShippingDetails (SESSION("token"))

ack = UCase(resArray("ACK"))

If ack="SUCCESS" Then

	Set resArray2 = ConfirmPayment (SESSION("finalPaymentAmount"))
	ack2 = UCase(resArray2("ACK"))
		'response.write ack2
'response.end
	If ack2="SUCCESS" OR ack2 = "SUCCESSWITHWARNING" Then

	Status_Pagamento= 1
		a= Split(session("finaliza"), ",")

		For i = 0 to Ubound(a)
		'response.write a(i)
		'response.end
		b = SetComprador(a(i))
		'response.write b & "<BR>"
		Next
	else
	Status_Pagamento=0
	end if

end if
'RESPONSE.END

'Response.Write("Retorno: " & Retorno & "<br>")

'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================

numero_Ped 		= "ABF-" & session("Numero_Pedido")

SQL_Historico_Pedido = 	"INSERT INTO Pedidos_Historico " &_
						"   (Numero_Pedido " &_
						"   ,Status_Pagamento " &_
						"   ,Numero_Transacao " &_
						"   ,Codigo_Autorizacao " &_
						"   ,Codigo_Paypal " &_
						"   ,Retorno) " &_
						"VALUES " &_
						"   ('" & Limpar_Texto(session("Numero_Pedido")) & "' " &_
						"   ,'" & Limpar_Texto(Status_Pagamento) & "' " &_
						"   ,'" & Limpar_Texto(SESSION("token")) & " - " & Limpar_Texto(SESSION("PAYERID")) & "' " &_
						"   ,'" & Limpar_Texto(ack2) & "' " &_
						"   ,'" & Limpar_Texto(numero_Ped) & "' " &_
						"   ,'" & ack2 & "')"
'Response.Write("Retorno: " & SQL_Historico_Pedido & "<br>")

Set Rs_Historico_Pedido = Conexao.Execute(SQL_Historico_Pedido)

Conexao.Close
'response.end
%>

<script>
	window.location = "http://credenciamento.btsinforma.com.br/tickets/retorno_exibir.asp?pedido=<%=Limpar_Texto(session("Numero_Pedido"))%>&transacao=<%=Limpar_Texto(SESSION("token"))& " - " & Limpar_Texto(SESSION("PAYERID"))%>"
	//window.location = "http://localhost:81/tickets/retorno_exibir.asp?pedido=<%=Limpar_Texto(session("Numero_Pedido"))%>&transacao=<%=Limpar_Texto(SESSION("token"))& " - " & Limpar_Texto(SESSION("PAYERID"))%>"

</script>
