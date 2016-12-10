<%
	Function ConsultarWS(CPF)
		
		Dim url_ws, urlws_wsdl    
		Dim obj_http, retorno
		Url = "https://www.mbxeventos.com/wsAOL/Methods.asmx"
		Usuario = "ABF"
		Senha = "RmhUrD7E"
		Metodo = "getXMLCPF"

    urlws = "https://www.mbxeventos.com/wsAOL/Methods.asmx" 'endereço do seu serviço
    urlws_wsdl = "https://www.mbxeventos.com/" 'note na imagem acima no header do xml que consta a entrada SOAPAction , ou seja, copiei fielmente conforma a imagem, note que LogOnASP é à ação ou método do ws.  

'abaixo monto o xml conforme a especificação.

    xml = "<?xml version =""1.0"" encoding=""UTF-8"" ?>" & vbCrLf
    xml = xml & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
    xml = xml & "<soap:Body>"
    xml = xml & "<"& Metodo &" xmlns=""http://tempuri.org/"">"
    xml = xml & "<user>ABF</user>"
    xml = xml & "<pass>RmhUrD7E</pass>"
    xml = xml & "<cpf>" & CPF & "</cpf>"
    xml = xml & " </" & Metodo & ">"
    xml = xml & "</soap:Body>"
    xml = xml & "</soap:Envelope>"

 
'abaixo segue a codificação utilizada para enviar o xml para o serviço.   
    SET obj_http = Server.CreateObject("Microsoft.XMLHTTP")
    obj_http.open "post", urlws, False
    obj_http.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    'obj_http.setRequestHeader "SOAPAction", urlws_wsdl
    obj_http.send xml
    retorno = obj_http.responseText

	'If ucase(registro.item(0).selectSingleNode("./"& verifica).text ) = ucase(CPF) Then
		
	'End If
	'teste = Mid(retorno,Instr(retorno,"CPF"),50)
	'Response.Write(Instr(retorno,"CPF"))
	'Response.Write teste
	Response.Write retorno.childNodes.item(0)
	Response.End
    SET OBJFSO = Nothing


    DIM OBJFSO, PASTA, ARQ

	End Function
%>
