<%
	Function ConsultarWS(CPF)
		Dim xmlhttp
		Dim objXMLDoc
		Dim Url, Usuario, Senha, Metodo
		Dim strSql
		'RESPONSE.WRITE("ConsultarWS")
		'RESPONSE.END
		Url = "https://www.mbxeventos.com/wsAOL/Methods.asmx"
		Usuario = "ABF"
		Senha = "RmhUrD7E"
		Metodo = "getXMLCPF"
		'strSend = "<?xml version=""1.0"" encoding=""utf-8""?>" &_
		'	"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
		'	"<soap:Body><" & Metodo & " xmlns=""http://tempuri.org/""><user>abfne</user><pass>xAfr9ojG</pass><cpf>" & CPF & "</cpf></" & Metodo & "></soap:Body>" &_
		'	"</soap:Envelope>"
		strSend = "<?xml version=""1.0"" encoding=""utf-8""?>" &_
			"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
			"<soap:Body><" & Metodo & " xmlns=""http://tempuri.org/""><user>ABF</user><pass>RmhUrD7E</pass><cpf>" & CPF & "</cpf></" & Metodo & "></soap:Body>" &_
			"</soap:Envelope>"
		
		Set xmlhttp = Server.CreateObject("MSXML2.XMLHTTP")
		xmlhttp.open "POST", Url, false
		xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=""utf-8"""
		xmlhttp.Send strSend
		xmlString = xmlhttp.responseText
		Set xmlhttp = Nothing
		
		Set objXMLDoc = CreateObject("Microsoft.XMLDOM") 
		objXMLDoc.async = False 
		objXMLDoc.LoadXml(xmlString)
		
		Set objXML = objXMLDoc.documentElement.selectSingleNode("/")
		objXMLDoc.LoadXml(objXML.Text)
		
		Set NodeList = objXML.getElementsByTagName("Table")
		'Response.Write 
		'Response.End
				
		For Each Node In NodeList
		
				'Response.Write  Node.childNodes(3).Text 
				'Response.Write  Node.childNodes(4).Text 
				'Response.Write  Node.childNodes(5).Text 
				'Response.Write  Node.childNodes(6).Text 
				'Response.Write  Node.childNodes(16).Text 
				'Response.Write  Session("cliente_edicao") 
				'Response.Write  Session("cliente_tipo")  
				'Response.Write  Session("cliente_idioma") 
				'Response.End
				
				'nascimento = Left(Node.childNodes(16).Text, 10)
				
				if InStr(Node.childNodes(16).Text,"/") then
					nascimento = Left(Node.childNodes(16).Text, 10)
				else
					nascimento = Left(Node.childNodes(18).Text, 10)
				end if
				
				'Response.Write nascimento
				'Response.End
		
			If Node.childNodes(3).Text = CPF Then
				strSql = "SET DATEFORMAT YDM; EXEC dbo.SP_IN_VISITANTES_CADASTRO"
				strSql = strSql & " @CPF = '" & Node.childNodes(3).Text & "'"
				strSql = strSql & " , @Email = '" & Node.childNodes(4).Text & "'"
				strSql = strSql & " , @Nome_Completo = '" & Node.childNodes(5).Text & "'"
				strSql = strSql & " , @Nome_Credencial = '" & Node.childNodes(6).Text & "'"
				strSql = strSql & " , @Data_Nasc = '" & nascimento & "'"
				strSql = strSql & " , @Id_Edicao = '" & Session("cliente_edicao") & "'"
				strSql = strSql & " , @Id_Tipo_Credenciamento = '" & Session("cliente_tipo") & "'"
				strSql = strSql & " , @Id_Idioma = '" & Session("cliente_idioma") & "'"
				
				Exit For
			End If
		Next
		
		If strSql <> "" Then
			Set objRS = Conexao.Execute(strSql)
			If Not objRS.EOF Then
				ConsultarWS = objRS("Id_Visitante")
			End If
		End If
		
		Set objRS = Nothing
		Set NodeList = Nothing
		Set objXML = Nothing
		Set objXMLDoc = Nothing
	End Function
	
	'Function EsCortesia()
	'	Dim xmlhttp
	'	Dim objXMLDoc
	'	Dim Url, Usuario, Senha, Metodo
	'	Dim strSql
	'	
	'	Url = "https://www.mbxeventos.com/wsAOL/Methods.asmx"
	'	Usuario = "ABF"
	'	Senha = "RmhUrD7E"
	'	Metodo = "EsCortesia"
	'	strSend = "<?xml version=""1.0"" encoding=""utf-8""?>" &_
	'		"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
	'		"<soap:Body><" & Metodo & " xmlns=""http://tempuri.org/""><user>abfne</user><pass>xAfr9ojG</pass><cpf>" & CPF & "</cpf></" & Metodo & "></soap:Body>" &_
	'		"</soap:Envelope>"
	'		
	'	'strSend = "<?xml version=""1.0"" encoding=""utf-8""?><soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:tem=""http://tempuri.org/""> <soapenv:Header/> <soapenv:Body> <tem:" & Metodo & ">       <tem:user>" & Usuario & "</tem:user>   <tem:pass>" & Senha & "</tem:pass>   "
    '
	'	if instr(CPF, "@") > 0  then
	'	strSend = strSend & "<tem:CPF></tem:CPF><tem:Email>" & CPF & "</tem:Email>"
	'	ELSE
	'	strSend = strSend & "<tem:CPF>" & CPF & "</tem:CPF><tem:Email></tem:Email>"
	'	END IF
    '
	'	strSend = strSend & " </tem:" & Metodo & ">   </soapenv:Body> </soapenv:Envelope>"	
	'		
	'	'response.write strSend
	'	'response.end
	'	Set xmlhttp = Server.CreateObject("MSXML2.XMLHTTP")
	'	xmlhttp.open "POST", Url, false
	'	xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=""utf-8"""
	'	xmlhttp.Send strSend
	'	xmlString = xmlhttp.responseText
	'	Set xmlhttp = Nothing
	'	'response.write xmlString
	'	'response.end
	'	Set objXMLDoc = CreateObject("Microsoft.XMLDOM") 
	'	objXMLDoc.async = False 
	'	objXMLDoc.LoadXml(xmlString)
	'	'objXMLDoc.LoadXml("<?xml version=""1.0"" encoding=""ISO-8859-1"" ?>	<Internet>		<Opcoes>			WEB - E-MAIL - VOZ				<Locaweb>					<Opcao>Hospedagem de Sites</Opcao>				</Locaweb>				<LocaMail>					<Opcao>Solucao para E-mails</Opcao>				</LocaMail>				<LocaVoz>					<Opcao>Portal de Voz</Opcao>				</LocaVoz>		</Opcoes>	</Internet>")		
	'	
	'	Set raiz = objXMLDoc.documentElement
    '
	'	if raiz.childNodes.item(0).childNodes.item(0).childNodes.item(0).text <> "" then
	'	
	'			ConsultarCortesia = raiz.childNodes.item(0).childNodes.item(0).childNodes.item(0).text
	'		
	'	End If
	'	
	'	Set objRS = Nothing
	'	Set NodeList = Nothing
	'	Set objXML = Nothing
	'	Set objXMLDoc = Nothing
	'End Function
	
	'Function ConsultarCortesia(CPF)
	'	Dim xmlhttp
	'	Dim objXMLDoc
	'	Dim Url, Usuario, Senha, Metodo
	'	Dim strSql
	'	
	'	Url = "https://www.mbxeventos.com/wsAOL/Methods.asmx"
	'	Usuario = "ABF"
	'	Senha = "RmhUrD7E"
	'	Metodo = "EsCortesia"
	'	strSend = "<?xml version=""1.0"" encoding=""utf-8""?>" &_
	'		"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
	'		"<soap:Body><" & Metodo & " xmlns=""http://tempuri.org/""><user>"& Usuario &"</user><pass>"& Senha &"</pass><cpf>" & CPF & "</cpf></" & Metodo & "></soap:Body>" &_
	'		"</soap:Envelope>"
	'		
	'	'strSend = "<?xml version=""1.0"" encoding=""utf-8""?><soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:tem=""http://tempuri.org/""> <soapenv:Header/> <soapenv:Body> <tem:" & Metodo & ">       <tem:user>" & Usuario & "</tem:user>   <tem:pass>" & Senha & "</tem:pass>   "
    '
	'	if instr(CPF, "@") > 0  then
	'	strSend = strSend & "<tem:CPF></tem:CPF><tem:Email>" & CPF & "</tem:Email>"
	'	ELSE
	'	strSend = strSend & "<tem:CPF>" & CPF & "</tem:CPF><tem:Email></tem:Email>"
	'	END IF
    '
	'	strSend = strSend & " </tem:" & Metodo & ">   </soapenv:Body> </soapenv:Envelope>"	
	'		
	'	'response.write strSend
	'	'response.end
	'	Set xmlhttp = Server.CreateObject("MSXML2.XMLHTTP")
	'	xmlhttp.open "POST", Url, false
	'	xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=""utf-8"""
	'	xmlhttp.Send strSend
	'	xmlString = xmlhttp.responseText
	'	Set xmlhttp = Nothing
	'	'response.write xmlString
	'	'response.end
	'	Set objXMLDoc = CreateObject("Microsoft.XMLDOM") 
	'	objXMLDoc.async = False 
	'	objXMLDoc.LoadXml(xmlString)
	'	'objXMLDoc.LoadXml("<?xml version=""1.0"" encoding=""ISO-8859-1"" ?>	<Internet>		<Opcoes>			WEB - E-MAIL - VOZ				<Locaweb>					<Opcao>Hospedagem de Sites</Opcao>				</Locaweb>				<LocaMail>					<Opcao>Solucao para E-mails</Opcao>				</LocaMail>				<LocaVoz>					<Opcao>Portal de Voz</Opcao>				</LocaVoz>		</Opcoes>	</Internet>")		
	'	
	'	Set raiz = objXMLDoc.documentElement
	'	
    '
	'	if raiz.childNodes.item(0).childNodes.item(0).childNodes.item(0).text <> "" then
	'	
	'			ConsultarCortesia = raiz.childNodes.item(0).childNodes.item(0).childNodes.item(0).text
	'		
	'	End If
	'	
	'	Set objRS = Nothing
	'	Set NodeList = Nothing
	'	Set objXML = Nothing
	'	Set objXMLDoc = Nothing
	'End Function
	
	Function SetCompradorOLD(CPF)
		Dim xmlhttp
		Dim objXMLDoc
		Dim Url, Usuario, Senha, Metodo
		Dim strSql
		
		Url = "https://www.mbxeventos.com/wsAOL/Methods.asmx"
		Usuario = "BTS"
		Senha = "RmhUrD7E"
		Metodo = "setComprador"
		strSend = "<?xml version=""1.0"" encoding=""utf-8""?>" &_
			"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
			"<soap:Body><" & Metodo & " xmlns=""http://tempuri.org/""><user>BTS</user><pass>RmhUrD7E</pass><cpf>" & CPF & "</cpf></" & Metodo & "></soap:Body>" &_
			"</soap:Envelope>"
			
		strSend = "<?xml version=""1.0"" encoding=""utf-8""?><soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:tem=""http://tempuri.org/""> <soapenv:Header/> <soapenv:Body> <tem:" & Metodo & ">       <tem:user>" & Usuario & "</tem:user>   <tem:pass>" & Senha & "</tem:pass>   "

		
			strSend = " <soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:tem=""http://tempuri.org/""> "
   strSend = strSend & " <soapenv:Header/> "
   strSend = strSend & " <soapenv:Body> "
      strSend = strSend & " <tem:" & Metodo & " > "
         
    'strSend = strSend & " <tem:user>ABF</tem:user> "
	strSend = strSend & " <tem:user>BTS</tem:user> "
         
         strSend = strSend & " <tem:pass>RmhUrD7E</tem:pass> "


		
		
		if instr(CPF, "@") > 0  then
		strSend = strSend & "<tem:CPF></tem:CPF><tem:Email>" & CPF & "</tem:Email>"
		ELSE
		strSend = strSend & "<tem:CPF>" & CPF & "</tem:CPF><tem:Email></tem:Email>"
		END IF
		
		
		
		strSend = strSend & " </tem:" & Metodo & ">   </soapenv:Body> </soapenv:Envelope>"	
			
		'response.write strSend
		'response.end
		Set xmlhttp = Server.CreateObject("MSXML2.XMLHTTP")
		xmlhttp.open "POST", Url, false
		xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=""utf-8"""
		xmlhttp.Send strSend
		xmlString = xmlhttp.responseText
		Set xmlhttp = Nothing
		'response.write xmlString
		'response.end
		Set objXMLDoc = CreateObject("Microsoft.XMLDOM") 
		objXMLDoc.async = False 
		objXMLDoc.LoadXml(xmlString)
		'objXMLDoc.LoadXml("<?xml version=""1.0"" encoding=""ISO-8859-1"" ?>	<Internet>		<Opcoes>			WEB - E-MAIL - VOZ				<Locaweb>					<Opcao>Hospedagem de Sites</Opcao>				</Locaweb>				<LocaMail>					<Opcao>Solucao para E-mails</Opcao>				</LocaMail>				<LocaVoz>					<Opcao>Portal de Voz</Opcao>				</LocaVoz>		</Opcoes>	</Internet>")		
		
		Set raiz = objXMLDoc.documentElement
 
'Looping para percorrer todos os elementos filhos

 
' response.end
		'Set objXML = objXMLDoc.documentElement.selectSingleNode("/")
		'objXMLDoc.LoadXml(objXML.Text)
		'response.write raiz.childNodes.item(0).childNodes.item(0).childNodes.item(0).text
		'response.end
		'response.write raiz.childNodes.item(0).childNodes.item(0).childNodes.item(0).text & "<br>2"
				'response.end
		if raiz.childNodes.item(0).childNodes.item(0).childNodes.item(0).text <> "" then
		
				SetComprador = raiz.childNodes.item(0).childNodes.item(0).childNodes.item(0).text
				
			
		End If
		
		Set objRS = Nothing
		Set NodeList = Nothing
		Set objXML = Nothing
		Set objXMLDoc = Nothing
	End Function
	
	Function SetComprador(CPF)
		Set Conexao = Server.CreateObject("ADODB.Connection")
		Conexao.Open Application("cnn")
		
		SQL_Email = " Select " &_
				" Email " &_
				" From Visitantes V " &_
				" JOIN Relacionamento_Cadastro RC ON RC.ID_Visitante = V.ID_Visitante " &_
				" Where  " &_
				"	V.CPF = '"& CPF &"'"&_
				"	AND RC.ID_Edicao = " & Session("cliente_edicao")
		Set RS_Email = Server.CreateObject("ADODB.Recordset")
		RS_Email.Open SQL_Email, Conexao, 3, 3
		
		If not RS_Email.EOF then
			Email = RS_Email("Email")
		End If
		
		Url = "https://www.mbxeventos.com/wsAOL/Methods.asmx"
		Usuario = "BTS"
		Senha = "RmhUrD7E"
		Metodo = "setComprador"
		strSend = "<?xml version=""1.0"" encoding=""utf-8""?>" &_
			"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
			"<soap:Body><" & Metodo & " xmlns=""http://tempuri.org/""><user>ABF</user><pass>RmhUrD7E</pass><CPF>" & CPF & "</CPF><email>" & Email &"</email></" & Metodo & "></soap:Body>" &_
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
		'Response.Write(xmlhttp.responseText)
		
	End Function
	
%>
