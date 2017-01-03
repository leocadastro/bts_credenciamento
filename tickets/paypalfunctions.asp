
<%
	' ===================================================
	' PayPal API Include file
	'
	' Defines all the global variables and the wrapper functions
	'-----------------------------------------------------------

	Dim gv_APIEndpoint
	Dim gv_APIUserName
	Dim gv_APIPassword
	Dim gv_APISignature
	Dim gv_Version
	Dim gv_BNCode

	Dim gv_ProxyServer
	Dim gv_ProxyServerPort
	Dim gv_Proxy
'	Dim resArray

	'----------------------------------------------------------------------------------
	' Authentication Credentials for making the call to the server
	'----------------------------------------------------------------------------------
	SandboxFlag = false
	if session("teste_paypal") = true then SandboxFlag = true

	'------------------------------------
	' PayPal API Credentials
	' Replace <API_USERNAME> with your API Username
	' Replace <API_PASSWORD> with your API Password
	' Replace <API_SIGNATURE> with your Signature
	'------------------------------------
	if SandboxFlag = true Then
		gv_APIUserName	= "cobrancabts-facilitator_api1.informa.com"
		gv_APIPassword	= "EVNBNNE2XRSPPH23"
		gv_APISignature = "A0AEeAwm9QxCtd0Tay2IFrhlDob.AcqaT9fpMpU8ut04zhdrli2yUvYW"
	else
		gv_APIUserName	= "cobrancabts_api1.informa.com"
		gv_APIPassword	= "RBLXS8L328MEZF3B"
		gv_APISignature = "AiPC9BjkCyDFQXbSkoZcgqH3hpacAIJZ5qxA3J62bQrpqmBfPmqmT5ES"

	end if

	'-----------------------------------------------------
	' The BN Code only applicable for partners
	'----------------------------------------------------
	gv_BNCode = "PP-ECWizard"

	'----------------------------------------------------------------------
	' Define the PayPal URLs.
	' 	This is the URL that the buyer is first sent to do authorize payment with their paypal account
	' 	change the URL depending if you are testing on the sandbox
	' 	or going to the live PayPal site
	'
	' For the sandbox, the URL is       https://www.sandbox.paypal.com/webscr&cmd=_express-checkout&token=
	' For the live site, the URL is        https://www.paypal.com/webscr&cmd=_express-checkout&token=
	'------------------------------------------------------------------------
	if SandboxFlag = true Then
		gv_APIEndpoint = "https://api-3t.sandbox.paypal.com/nvp"
		PAYPAL_URL = "https://www.sandbox.paypal.com/webscr?cmd=_express-checkout&token="
	Else
		gv_APIEndpoint = "https://api-3t.paypal.com/nvp"
		PAYPAL_URL = "https://www.paypal.com/cgi-bin/webscr?cmd=_express-checkout&token="
	End If

	gv_Version	= "93"

	'WinObjHttp Request proxy settings.
	gv_ProxyServer	= "127.0.0.1"
	gv_ProxyServerPort = "808"
	gv_Proxy		= 2	'setting for proxy activation
	gv_UseProxy		= False

	'-------------------------------------------------------------------------------------------------------------------------------------------
	' Purpose: 	Prepares the parameters for the SetExpressCheckout API Call.
	' Inputs:
	'		paymentAmount:  	Total value of the shopping cart
	'		currencyCodeType: 	Currency code value the PayPal API
	'		paymentType: 		paymentType has to be one of the following values: Sale or Order or Authorization
	'		returnURL:			the page where buyers return to after they are done with the payment review on PayPal
	'		cancelURL:			the page where buyers return to when they cancel the payment	 review on PayPal
	' Returns:
	'		The NVP Collection object of the SetExpressCheckout call Response.
	'--------------------------------------------------------------------------------------------------------------------------------------------
	Function CallShortcutExpressCheckout( paymentAmount, currencyCodeType, paymentType, returnURL, cancelURL)

	'paymentAmount = 60
		'------------------------------------------------------------------------------------------------------------------------------------
		' Construct the parameter string that describes the SetExpressCheckout API call in the shortcut implementation
		'------------------------------------------------------------------------------------------------------------------------------------

		nvpstr	= "&" & Server.URLEncode("PAYMENTREQUEST_0_AMT") & "=" & Server.URLEncode(paymentAmount) & _
				  "&" & Server.URLEncode("PAYMENTREQUEST_0_PAYMENTACTION")&"=" & Server.URLEncode(paymentType) & _
				  "&"& Server.URLEncode("RETURNURL") & "=" & Server.URLEncode(returnURL) & _
				  "&" & Server.URLEncode("CANCELURL") & "=" & Server.URLEncode(cancelURL) & _
				  "&"& server.UrlEncode("PAYMENTREQUEST_0_CURRENCYCODE") & "=" & Server.URLEncode(currencyCodeType)


  '==================================================
  Set Conexao = Server.CreateObject("ADODB.Connection")
  Conexao.Open Application("cnn")
  '==================================================

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

  	Valor_Pedido	= FormatNumber(RS_Consulta_Pedidos("Valor_Pedido"),2)

  End If

	RS_Consulta_Pedidos.Close

' Session finaliza está trazendo o CPF e a Session cpf não está trazendo nada - Andre Alves
a= Split(session("finaliza"), ",")
Session("cpf") = session("finaliza")
 'Response.Write Session("cpf")
 'Response.End
For i = 0 to Ubound(a)
'Response.Write "Entrou"
'Response.End
	nvpstr	= nvpstr & "&" & Server.URLEncode("L_PAYMENTREQUEST_0_NAME"&i) & "=" & Server.URLEncode("Convite ABF")
	if instr(a(i),"@") > 0 then
		nvpstr	= nvpstr & "&" & Server.URLEncode("L_PAYMENTREQUEST_0_DESC"&i) & "=" & Server.URLEncode("Convite para E-mail: " & a(i))
	else
		nvpstr	= nvpstr & "&" & Server.URLEncode("L_PAYMENTREQUEST_0_DESC"&i) & "=" & Server.URLEncode("Convite para CPF:" & a(i))
	end if
	'Local de troca de valores de compra no PayPal
	nvpstr	= nvpstr & "&" & Server.URLEncode("L_PAYMENTREQUEST_0_AMT"&i) & "=" & Server.URLEncode(Replace(Valor_Pedido,",","."))

	'Response.Write(nvpstr)
	'Response.End

	nvpstr	= nvpstr & "&" & Server.URLEncode("L_PAYMENTREQUEST_0_QTY"&i) & "=" & Server.URLEncode(1)
Next

Conexao.Close

if len(Session("cpf")) = 11  then
'Response.Write "CPF"
'Response.End
					nvpstr = nvpstr & "&" & Server.URLEncode("PAYMENTREQUEST_0_SHIPTONAME") & "=" & Server.URLEncode(session("nome"))
					nvpstr = nvpstr &  "&" & Server.URLEncode("TAXIDTYPE") & "=" & Server.URLEncode("BR_CPF")
					nvpstr = nvpstr &  "&" & Server.URLEncode("TAXID") & "=" & Server.URLEncode(session("cpf"))
					nvpstr = nvpstr &  "&" & Server.URLEncode("EMAIL") & "=" & Server.URLEncode(session("email"))
				  else
				  nvpstr = nvpstr &  "&" & Server.URLEncode("PAYMENTREQUEST_0_SHIPTONAME") & "=" & Server.URLEncode(session("nome"))
				  nvpstr = nvpstr &  "&" & Server.URLEncode("EMAIL") & "=" & Server.URLEncode(session("email"))
				  end if



		SESSION("currencyCodeType")	= currencyCodeType
		SESSION("PaymentType")	= paymentType

		'---------------------------------------------------------------------------------------------------------------
		' Make the API call to PayPal
		' If the API call succeded, then redirect the buyer to PayPal to begin to authorize payment.
		' If an error occured, show the resulting errors
		'---------------------------------------------------------------------------------------------------------------

		Set resArray = hash_call("SetExpressCheckout",nvpstr)


'Response.Write resArray("L_ERRORCODE0")
'Response.Write resArray("L_SHORTMESSAGE0")
'Response.Write("\n")
'Response.Write resArray("L_LONGMESSAGE0")
'Response.End
		ack = UCase(resArray("ACK"))
		If ack="SUCCESS" Then
			' Save the token parameter in the Session
			SESSION("token") = resArray("TOKEN")
		End If

		set CallShortcutExpressCheckout	= resArray

	End Function

	'-------------------------------------------------------------------------------------------------------------------------------------------
	' Purpose: 	Prepares the parameters for the SetExpressCheckout API Call.
	' Inputs:
	'		paymentAmount:  	Total value of the shopping cart
	'		currencyCodeType: 	Currency code value the PayPal API
	'		paymentType: 		paymentType has to be one of the following values: Sale or Order or Authorization
	'		returnURL:			the page where buyers return to after they are done with the payment review on PayPal
	'		cancelURL:			the page where buyers return to when they cancel the payment review on PayPal
	'		shipToName:		the Ship to name entered on the merchant's site
	'		shipToStreet:		the Ship to Street entered on the merchant's site
	'		shipToCity:			the Ship to City entered on the merchant's site
	'		shipToState:		the Ship to State entered on the merchant's site
	'		shipToCountryCode:	the Code for Ship to Country entered on the merchant's site
	'		shipToZip:			the Ship to ZipCode entered on the merchant's site
	'		shipToStreet2:		the Ship to Street2 entered on the merchant's site
	'		phoneNum:			the phoneNum  entered on the merchant's site
	' Returns:
	'		The NVP Collection object of the SetExpressCheckout call Response.
	'--------------------------------------------------------------------------------------------------------------------------------------------
	Function CallMarkExpressCheckout(paymentAmount, currencyCodeType, paymentType, returnURL, cancelURL, shipToName, shipToStreet, shipToCity, shipToState, shipToCountryCode, shipToZip, shipToStreet2, phoneNum)
		'------------------------------------------------------------------------------------------------------------------------------------
		' Construct the parameter string that describes the SetExpressCheckout API call in the shortcut implementation
		'------------------------------------------------------------------------------------------------------------------------------------

		nvpstr	= "&" & Server.URLEncode("PAYMENTREQUEST_0_AMT") & "=" & Server.URLEncode(paymentAmount) & _
				  "&" & Server.URLEncode("PAYMENTREQUEST_0_PAYMENTACTION")&"=" & Server.URLEncode(paymentType) & _
				  "&" & Server.URLEncode("RETURNURL") & "=" & Server.URLEncode(returnURL) & _
				  "&" & Server.URLEncode("CANCELURL") & "=" & Server.URLEncode(cancelURL) & _
				  "&" & Server.URLEncode("ADDROVERRIDE") & "=1" & _
				  "&" & Server.URLEncode("PAYMENTREQUEST_0_SHIPTONAME") & "=" & Server.URLEncode(shipToName) & _
				  "&" & Server.URLEncode("PAYMENTREQUEST_0_SHIPTOSTREET") & "=" & Server.URLEncode(shipToStreet) & _
				  "&" & Server.URLEncode("PAYMENTREQUEST_0_SHIPTOSTREET2") & "=" & Server.URLEncode(shipToStreet2) & _
				  "&" & Server.URLEncode("PAYMENTREQUEST_0_SHIPTOCITY") & "=" & Server.URLEncode(shipToCity) & _
				  "&" & Server.URLEncode("PAYMENTREQUEST_0_SHIPTOSTATE") & "=" & Server.URLEncode(shipToState) & _
				  "&" & Server.URLEncode("PAYMENTREQUEST_0_SHIPTOCOUNTRYCODE") & "=" & Server.URLEncode(shipToCountryCode) & _
				  "&" & Server.URLEncode("PAYMENTREQUEST_0_SHIPTOZIP") & "=" & Server.URLEncode(shipToZip) & _
				  "&" & Server.URLEncode("PAYMENTREQUEST_0_SHIPTOPHONENUM") & "=" & Server.URLEncode(phoneNum) & _
				  "&"& server.UrlEncode("PAYMENTREQUEST_0_CURRENCYCODE") & "=" & Server.URLEncode(currencyCodeType)

  	    SESSION("currencyCodeType")	= currencyCodeType
		SESSION("PaymentType")	= paymentType

		'---------------------------------------------------------------------------
		' Make the API call to PayPal to set the Express Checkout token
		' 	If the API call succeded, then redirect the buyer to PayPal to begin to authorize payment.
		' 	If an error occured, show the resulting errors
		'---------------------------------------------------------------------------
		Set resArray = hash_call("SetExpressCheckout",nvpstr)

		ack = UCase(resArray("ACK"))
		If ack="SUCCESS" Then
			' Save the token parameter in the Session
			SESSION("token") = resArray("TOKEN")
		End If

		set CallMarkExpressCheckout	= resArray

	End Function

	'-------------------------------------------------------------------------------------------------------------------------------------------
	' Purpose: 	Prepares the parameters for the GetExpressCheckoutDetails API and makes the API call.
	'
	' Inputs:
	'		token: 	The token value returned by the SetExpressCheckout call
	' Returns:
	'		The NVP Collection object of the GetExpressCheckoutDetails Call Response.
	'--------------------------------------------------------------------------------------------------------------------------------------------
	Function GetShippingDetails( token )
		'---------------------------------------------------------------------------
		' At this point, the buyer has completed authorizing the payment
		' at PayPal.  The function will call PayPal to obtain the details
		' of the authorization, incuding any shipping information of the
		' buyer.  Remember, the authorization is not a completed transaction
		' at this state - the buyer still needs an additional step to finalize
		' the transaction
		'---------------------------------------------------------------------------

	    '---------------------------------------------------------------------------
		' Build a second API request to PayPal, using the token as the
		'  ID to get the details on the payment authorization
		'---------------------------------------------------------------------------
		nvpstr="&TOKEN=" & token

		'---------------------------------------------------------------------------
		' Make the API call and store the results in an array.
		'	If the call was a success, show the authorization details, and provide
		' 	an action to complete the payment.
		'	If failed, show the error
		'---------------------------------------------------------------------------
		set resArray = hash_call("GetExpressCheckoutDetails",nvpstr)
		ack = UCase(resArray("ACK"))
		If ack="SUCCESS" Then
			' Save the token parameter in the Session
			SESSION("PAYERID") = resArray("PAYERID")
			'response.write resArray("PAYERID")
		End If
		set GetShippingDetails = resArray
	End Function

	'-------------------------------------------------------------------------------------------------------------------------------------------
	' Purpose: 	Prepares the parameters for the GetExpressCheckoutDetails API and makes the call.
	'
	' Inputs:
	'		finalPaymentAmount:  	The final total of the shopping cart including Shipping, Handling and other fees
	' Returns:
	'		The NVP Collection object of the DoExpressCheckoutPayment Call Response.
	'--------------------------------------------------------------------------------------------------------------------------------------------
	Function ConfirmPayment( finalPaymentAmount )

		'------------------------------------------------------------------------------------------------------------------------------------
		'----	Use the values stored in the session from the previous SetEC call
		'------------------------------------------------------------------------------------------------------------------------------------
		token			= SESSION("token")
		currCodeType	= SESSION("currencyCodeType")
		paymentType		= SESSION("PaymentType")
		payerID			= SESSION("PayerID")

		nvpstr			=	"&" & Server.URLEncode("TOKEN") & "=" & Server.URLEncode(token) & "&" &_
							Server.URLEncode("PAYERID")&"=" &Server.URLEncode(payerID) & "&" &_
							Server.URLEncode("PAYMENTREQUEST_0_PAYMENTACTION")&"=" & Server.URLEncode(paymentType) & "&" &_
							Server.URLEncode("PAYMENTREQUEST_0_AMT") &"=" & Server.URLEncode(finalPaymentAmount) & "&" &_
							Server.URLEncode("PAYMENTREQUEST_0_CURRENCYCODE")& "=" &Server.URLEncode(currCodeType)
		'-------------------------------------------------------------------------------------------
		' Make the call to PayPal to finalize payment
		' If an error occured, show the resulting errors
		'-------------------------------------------------------------------------------------------
		set ConfirmPayment = hash_call("DoExpressCheckoutPayment",nvpstr)
	End Function

	'-------------------------------------------------------------------------------------------------------------------------------------------
	' Purpose: 	Prepares the parameters for the DoDirectPayment API and makes the call.
	'
	' Inputs:
	'		paymentType: 		paymentType has to be one of the following values: Sale or Order or Authorization
	'		paymentAmount:  		Total value of the shopping cart
	'		creditCardType		Credit card type has to one of the following values: Visa or MasterCard or Discover or Amex or Switch or Solo
	'		creditCardNumber	Credit card number
	'		expDate			Credit expiration date
	'		cvv2				CVV2
	'		firstName			Customer's First Name
	'		lastName			Customer's Last Name
	'		street			Customer's Street Address
	'		city				Customer's City
	'		state				Customer's State
	'		zip				Customer's Zip
	'		countryCode		Customer's Country represented as a PayPal CountryCode
	'		currencyCode		Customer's Currency represented as a PayPal CurrencyCode
	'
	' Returns:
	'		The NVP Collection object of the DoDirectPayment Call Response.
	'--------------------------------------------------------------------------------------------------------------------------------------------
	Function DirectPayment( paymentType, paymentAmount, creditCardType, creditCardNumber, expDate, cvv2, firstName, lastName, street, city, state, zip, countryCode, currencyCode )

		' Construct the parameter string that describes the SetExpressCheckout API call in the shortcut implementation

		nvpstr	=	"&PAYMENTACTION=" & paymentType & _
					"&AMT=" & paymentAmount &_
					"&CREDITCARDTYPE=" & creditCardType &_
					"&ACCT=" & creditCardNumber & _
					"&EXPDATE=" & expDate &_
					"&CVV2=" & cvv2 &_
					"&FIRSTNAME=" & firstName &_
					"&LASTNAME=" & lastName &_
					"&STREET=" & street &_
					"&CITY=" & city &_
					"&STATE=" & state &_
					"&ZIP=" & zip &_
					"&COUNTRYCODE=" & countryCode &_
					"&CURRENCYCODE=" & currencyCode

		nvpstr	=	URLEncode(nvpstr)

		'-------------------------------------------------------------------------------------------
		' Make the call to PayPal to finalize payment
		' If an error occured, show the resulting errors
		'-------------------------------------------------------------------------------------------
		set DirectPayment = hash_call("DoDirectPayment",nvpstr)
	End Function

	'----------------------------------------------------------------------------------
	' Purpose: 	Make the API call to PayPal, using API signature.
	' Inputs:
	'		Method name to be called & NVP string to be sent with the post method
	' Returns:
	'		NVP Collection object of Call Response.
	'----------------------------------------------------------------------------------
	Function hash_call ( methodName,nvpStr )
		Set objHttp = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")

		nvpStrComplete	= "METHOD=" & Server.URLEncode(methodName) & "&VERSION=" & Server.URLEncode(gv_Version) & "&USER=" & Server.URLEncode(gv_APIUserName) & "&PWD=" & Server.URLEncode(gv_APIPassword) & "&SIGNATURE=" & Server.URLEncode(gv_APISignature) & nvpStr
		nvpStrComplete	= nvpStrComplete & "&BUTTONSOURCE=" & Server.URLEncode( gv_BNCode )

		Set SESSION("nvpReqArray")= deformatNVP( nvpStrComplete )
		'objHttp.open "POST", gv_APIEndpoint, False
		'WinHttpRequestOption_SslErrorIgnoreFlags=4
		'objHttp.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300

		Set xmlhttp = Server.CreateObject("MSXML2.XMLHTTP")
		xmlhttp.open "POST", gv_APIEndpoint, false
		xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=""utf-8"""
		xmlhttp.Send nvpStrComplete
		xmlString = xmlhttp.responseText

		'R''esponse.Write xmlString
		'Response.End
		'Não está entrando no if abaixo - Andre Alves

		'response.write nvpStrComplete
		'response.end
		'objHttp.Send nvpStrComplete
		'response.write 1
		'response.end
		Set nvpResponseCollection =deformatNVP(xmlString)


		Set hash_call = nvpResponseCollection
		Set objHttp = Nothing

		If Err.Number <> 0 Then
			SESSION("Message")	= ErrorFormatter(Err.Description,Err.Number,Err.Source,"hash_call")
			SESSION("nvpReqArray") =  Null
		Else
			SESSION("Message")	= Null
		End If

	End Function

	'----------------------------------------------------------------------------------
	' Purpose: 	Formats the error Messages.
	' Inputs:
	'
	' Returns:
	'		Formatted Error string
	'----------------------------------------------------------------------------------
	Function ErrorFormatter ( errDesc, errNumber, errSource, errlocation )
		ErrorFormatter ="<font color=red>" & _
								"<TABLE align = left>" &_
								"<TR>" &"<u>Error Occured!!!</u>" & "</TR>" &_
								"<TR>" &"<TD>Error Description :</TD>" &"<TD>"&errDesc& "</TD>"& "</TR>" &_
								"<TR>" &"<TD>Error number :</TD>" &"<TD>"&errNumber& "</TD>"& "</TR>" &_
								"<TR>" &"<TD>Error Source :</TD>" &"<TD>"&errSource& "</TD>"& "</TR>" &_
								"<TR>" &"<TD>Error Location :</TD>" &"<TD>"&errlocation& "</TD>"& "</TR>" &_
								"</TABLE>" &_
								"</font>"
	End Function

	'----------------------------------------------------------------------------------
	' Purpose: 	Convert nvp string to Collection object.
	' Inputs:
	'		NVP string.
	' Returns:
	'		NVP Collection object created from deserializing the NVP string.
	'----------------------------------------------------------------------------------
	Function deformatNVP ( nvpstr )
		On Error Resume Next

		Dim AndSplitedArray,EqualtoSplitedArray,Index1,Index2,NextIndex

		Set NvpCollection = Server.CreateObject("Scripting.Dictionary")
		AndSplitedArray = Split(nvpstr, "&", -1, 1)
		NextIndex=0
		'RESPONSE.WRITE AndSplitedArray(0)

		For Index1 = 0 To UBound(AndSplitedArray)
			EqualtoSplitedArray=Split(AndSplitedArray(Index1), "=", -1, 1)
			For Index2 = 0 To UBound(EqualtoSplitedArray)
			'RESPONSE.WRITE EqualtoSplitedArray(Index2) & " = " & EqualtoSplitedArray(Index2+1)
				NextIndex=Index2+1
				NvpCollection.Add URLDecode(EqualtoSplitedArray(Index2)),URLDecode(EqualtoSplitedArray(NextIndex))
				Index2=Index2+1
			Next
		Next

		'RESPONSE.WRITE NvpCollection("ACK")
		'RESPONSE.END
		Set deformatNVP = NvpCollection
		If Err.Number <> 0 Then
			SESSION("Message")	= ErrorFormatter(Err.Description,Err.Number,Err.Source,"deformatNVP")
		else
			SESSION("Message")	= Null
		End If
	End Function

	'----------------------------------------------------------------------------------
	' Purpose: URL Encodes a string
	' Inputs:
	'		String to be url encoded.
	' Returns:
	'		Url Encoded string.
	'----------------------------------------------------------------------------------
	Function URLEncode(str)
		On Error Resume Next

	    Dim AndSplitedArray,EqualtoSplitedArray,Index1,Index2,UrlEncodeString,NvpUrlEncodeString

		AndSplitedArray = Split(nvpstr, "&", -1, 1)
		UrlEncodeString=""
		NvpUrlEncodeString=""

		For Index1 = 0 To UBound(AndSplitedArray)
			EqualtoSplitedArray=Split(AndSplitedArray(Index1), "=", -1, 1)
			For Index2 = 0 To UBound(EqualtoSplitedArray)
			If Index2 = 0 then
				UrlEncodeString=UrlEncodeString & Server.URLEncode(EqualtoSplitedArray(Index2))
			Else
				UrlEncodeString=UrlEncodeString &"="& Server.URLEncode(EqualtoSplitedArray(Index2))
			End if
			Next
			If Index1 = 0 then
				NvpUrlEncodeString= NvpUrlEncodeString & UrlEncodeString
			Else
				NvpUrlEncodeString= NvpUrlEncodeString &"&"&UrlEncodeString
			End if
			UrlEncodeString=""
		Next
		URLEncode = NvpUrlEncodeString

		If Err.Number <> 0 Then
			SESSION("Message")	= ErrorFormatter(Err.Description,Err.Number,Err.Source,"URLEncode")
		else
			SESSION("Message")	= Null
		End If

	 End Function

	'----------------------------------------------------------------------------------
	' Purpose: Decodes a URL Encoded string
	' Inputs:
	'		A URL encoded string
	' Returns:
	'		Decoded string.
	'----------------------------------------------------------------------------------
	Function URLDecode(str)
		On Error Resume Next

		str = Replace(str, "+", " ")
		For i = 1 To Len(str)
			sT = Mid(str, i, 1)
			If sT = "%" Then
				If i+2 < Len(str) Then
					sR = sR & _
						Chr(CLng("&H" & Mid(str, i+1, 2)))
					i = i+2
				End If
			Else
				sR = sR & sT
			End If
		Next

		URLDecode = sR
		If Err.Number <> 0 Then
			SESSION("Message")	= ErrorFormatter(Err.Description,Err.Number,Err.Source,"URLDecode")
		else
			SESSION("Message")	= Null
		End If

	End Function

	'----------------------------------------------------------------------------------
	' Purpose: 	It's Workaround Method for Response.Redirect
	'          	It will redirect the page to the specified url without urlencoding
	' Inputs:
	'		Url to redirect the page
	'----------------------------------------------------------------------------------
	Function ReDirectURL( token )
		On Error Resume Next

		payPalURL = PAYPAL_URL & token
		response.clear
		response.status="302 Object moved"
		response.AddHeader "location", payPalURL
		If Err.Number <> 0 Then
			SESSION("Message")	= ErrorFormatter(Err.Description,Err.Number,Err.Source,"ReDirectURL")
		else
			SESSION("Message")	= Null
		End If
	End Function

%>
