<%

CPF = "12345678900"
		Url = "https://www.mbxeventos.com/wsAOL/Methods.asmx"
		Usuario = "BTS"
		Senha = "RmhUrD7E"
		Metodo = "setComprador"
		strSend = "<?xml version=""1.0"" encoding=""utf-8""?>" &_
			"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
			"<soap:Body><" & Metodo & " xmlns=""http://tempuri.org/""><user>ABF</user><pass>RmhUrD7E</pass><CPF>" & CPF & "</CPF><email>andre.alves@informa.com</email></" & Metodo & "></soap:Body>" &_
			"</soap:Envelope>"
   	Dim xmlhttp
    Dim DataToSend
    DataToSend=strSend
    Dim postUrl
    postUrl = Url
    Set xmlhttp = server.Createobject("MSXML2.XMLHTTP")
    xmlhttp.Open "POST",postUrl,false
    xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=""utf-8"""
    xmlhttp.send DataToSend
    Response.Write(xmlhttp.responseText)
	'Reponse.End
%>