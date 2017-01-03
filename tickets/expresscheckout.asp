<%@Language="VBSCRIPT"%>
<!-- #include file ="paypalfunctions.asp" -->
<%
' ==================================
' PayPal Express Checkout Module
' ==================================

On Error Resume Next

'------------------------------------
' The paymentAmount is the total value of
' the shopping cart, that was set
' earlier in a session variable
' by the shopping cart page
'------------------------------------
paymentAmount = REQUEST.form("ValorDocumento")
'Response.Write REQUEST.form("ValorDocumento")
'Response.End

session("finalPaymentAmount") = paymentAmount
'Response.Write session("finalPaymentAmount")
'Response.End

'------------------------------------
' The currencyCodeType and paymentType
' are set to the selections made on the Integration Assistant
'------------------------------------
currencyCodeType = "BRL"
paymentType = "Sale"

'------------------------------------
' The returnURL is the location where buyers return to when a
' payment has been succesfully authorized.
'
' This is set to the value entered on the Integration Assistant
'------------------------------------
'returnURL = "HTTP://CREdenciamento.btsinforma.com.br/tickets/retorno.asp"
returnURL = "http://localhost:81/tickets/retorno.asp"

'------------------------------------
' The cancelURL is the location buyers are sent to when they click the
' return to XXXX site where XXX is the merhcant store name
' during payment review on PayPal
'
' This is set to the value entered on the Integration Assistant
'------------------------------------
'cancelURL = "http://credenciamento.btsinforma.com.br/tickets/pagamento.asp"
cancelURL = "http://localhost:81/tickets/pagamento.asp"

'------------------------------------
' Calls the SetExpressCheckout API call
'
' The CallShortcutExpressCheckout function is defined in the file PayPalFunctions.asp,
' it is included at the top of this file.
'-------------------------------------------------
'Response.Write paymentAmount & "<br/>"
'Response.Write currencyCodeType & "<br/>"
'Response.Write paymentType & "<br/>"
'Response.Write returnURL & "<br/>"
'Response.Write cancelURL
'Response.End
Set resArray = CallShortcutExpressCheckout (paymentAmount, currencyCodeType, paymentType, returnURL, cancelURL)
'ack está retornando Failure - Andre Alves
ack = UCase(resArray("ACK"))

Response.Write ack
Response.end
If ack="SUCCESS" Then
	' Redirect to paypal.com
	ReDirectURL( resArray("TOKEN") )
Else
	'Display a user friendly Error on the page using any of the following error information returned by PayPal
	ErrorCode = URLDecode( resArray("L_ERRORCODE0"))
	ErrorShortMsg = URLDecode( resArray("L_SHORTMESSAGE0"))
	ErrorLongMsg = URLDecode( resArray("L_LONGMESSAGE0"))
	ErrorSeverityCode = URLDecode( resArray("L_SEVERITYCODE0"))
	'Caso a Sessão caia, Redirecionar para página
	'Response.Redirect ("/tickets/default.asp?dc="& Session("finaliza")&"&msg=expresscheckerro")
	'Response.Write ErrorSeverityCode
	'response.end
End If
%>
