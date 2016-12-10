<%@Language="VBSCRIPT"%>
<!-- #include file ="paypalfunctions.asp" -->
<%
' ==================================
' PayPal Express Checkout Module
' ==================================

On Error Resume Next
token = request("token")
'------------------------------------
' The paymentAmount is the total value of 
' the shopping cart, that was set 
' earlier in a session variable 
' by the shopping cart page
'------------------------------------
paymentAmount = REQUEST.form("ValorDocumento")

session("finalPaymentAmount") = paymentAmount

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
returnURL = "HTTP://CREdenciamento.btsinforma.com.br/tickets/retorno.asp"
'returnURL = "HTTP://ws.homologabts.com.br/tickets/retorno.asp"

'------------------------------------
' The cancelURL is the location buyers are sent to when they click the
' return to XXXX site where XXX is the merhcant store name
' during payment review on PayPal
'
' This is set to the value entered on the Integration Assistant 
'------------------------------------
cancelURL = "http://credenciamento.btsinforma.com.br/tickets/pagamento.asp"
'cancelURL = "http://ws.homologabts.com.br/tickets/pagamento.asp"

'------------------------------------
' Calls the SetExpressCheckout API call
'
' The CallShortcutExpressCheckout function is defined in the file PayPalFunctions.asp,
' it is included at the top of this file.
'-------------------------------------------------
'response.write token
Set resArray = GetShippingDetails (token)

response.write resArray

%>
