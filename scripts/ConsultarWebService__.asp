<%
	Function ConsultarWS(CPF)
		Dim xmlhttp
		Dim objXMLDoc
		Dim Url, Usuario, Senha, Metodo
		Dim strSql
		
		Url = "https://www.mbxeventos.com/wsAOL/Methods.asmx"
		Usuario = "ABF"
		Senha = "RmhUrD7E"
		Metodo = "getXMLCPF"
		
		if instr(CPF, "@") > 0  then Metodo = "getXMLEmail"
		
		
		
		strSend = "<?xml version=""1.0"" encoding=""utf-8""?>" &_
			"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
			"<soap:Body><" & Metodo & " xmlns=""http://tempuri.org/""><user>abfne</user><pass>xAfr9ojG</pass><cpf>" & CPF & "</cpf></" & Metodo & "></soap:Body>" &_
			"</soap:Envelope>"
			
			
		strSend = "<?xml version=""1.0"" encoding=""utf-8""?><soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:tem=""http://tempuri.org/""> <soapenv:Header/> <soapenv:Body> <tem:" & Metodo & ">       <tem:user>" & Usuario & "</tem:user>   <tem:pass>" & Senha & "</tem:pass> "
		
		if instr(CPF, "@") > 0  then
		strSend = strSend & "<tem:email>" & CPF & "</tem:email>    </tem:" & Metodo & ">   </soapenv:Body> </soapenv:Envelope>"	
		else
		
		strSend = strSend & "<tem:cpf>" & CPF & "</tem:cpf>    </tem:" & Metodo & ">   </soapenv:Body> </soapenv:Envelope>"	
		
		end if
			
		
		Set xmlhttp = Server.CreateObject("MSXML2.XMLHTTP")
		xmlhttp.open "POST", Url, false
		xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=""utf-8"""
		xmlhttp.Send strSend
		xmlString = xmlhttp.responseText
		'Response.Write xmlhttp.responseText
		'Response.End
		Set xmlhttp = Nothing
		
		Set objXMLDoc = CreateObject("Microsoft.XMLDOM") 
		objXMLDoc.async = False 
		objXMLDoc.LoadXml(xmlString)
		'objXMLDoc.LoadXml("<?xml version=""1.0"" encoding=""ISO-8859-1"" ?>	<Internet>		<Opcoes>			WEB - E-MAIL - VOZ				<Locaweb>					<Opcao>Hospedagem de Sites</Opcao>				</Locaweb>				<LocaMail>					<Opcao>Solucao para E-mails</Opcao>				</LocaMail>				<LocaVoz>					<Opcao>Portal de Voz</Opcao>				</LocaVoz>		</Opcoes>	</Internet>")		
		
		Set raiz = objXMLDoc.documentElement
		set registro = objXMLDoc.getElementsByTagName("Table")
		response.write raiz.selectNodes("//CPF")
		response.end
'Looping para percorrer todos os elementos filhos

'response.write registro.item(0).selectSingleNode("./Nascimento").text 
 'response.end
		'Set objXML = objXMLDoc.documentElement.selectSingleNode("/")
		'objXMLDoc.LoadXml(objXML.Text)
		'response.write raiz.childNodes.item(0).childNodes.item(0).childNodes.item(0).childNodes.item(0).text
		'response.end
		if raiz.childNodes.item(0).childNodes.item(0).childNodes.item(0).childNodes.item(0).text <> "" then
		'Set NodeList = raiz.childNodes.item(0).childNodes.item(0).childNodes.item(0).childNodes.item(0).childNodes.item(0)
		

'		For Each Node In NodeList

verifica = "CPF"
ident_data = 17
	if instr(CPF, "@") > 0  then 
		verifica = "email"
		ident_data=17
		end if

		'response.write raiz.childNodes.item(0).childNodes.item(0).childNodes.item(0).childNodes.item(0).text
		'response.end 
			If ucase(registro.item(0).selectSingleNode("./"& verifica).text ) = ucase(CPF) Then
		'response.write raiz.childNodes.item(0).childNodes.item(0).childNodes.item(0).childNodes.item(0).childNodes.item(0).childNodes.item(17).text
			'response.end
			data = Left(registro.item(0).selectSingleNode("./Nascimento").text, 10)
			
			ano = left(data,4)
			dia = right(data,2)
			mes = replace(data, ano&"-", "")
			mes= replace(mes, "-" & dia, "")
		'ucase(raiz.childNodes.item(0).childNodes.item(0).childNodes.item(0).childNodes.item(0).childNodes.item(0).childNodes.item(verifica).text)
		'	response.end
			if isnumeric(dia) and isnumeric(mes) and isnumeric(ano) then
			session("nome")= registro.item(0).selectSingleNode("./Nome").text & " " & registro.item(0).selectSingleNode("./Sobrenome").text
			session("cpf")= registro.item(0).selectSingleNode("./CPF").text
			session("email") = registro.item(0).selectSingleNode("./email").text
			
			
				strSql = "SET DATEFORMAT YDM; EXEC dbo.SP_IN_VISITANTES_CADASTRO"
				strSql = strSql & " @CPF = '" & registro.item(0).selectSingleNode("./CPF").text & "'"
				strSql = strSql & " , @Email = '" & registro.item(0).selectSingleNode("./email").text & "'"
				strSql = strSql & " , @Nome_Completo = '" & registro.item(0).selectSingleNode("./Nome").text & " " & registro.item(0).selectSingleNode("./Sobrenome").text & "'"
				strSql = strSql & " , @Nome_Credencial = '" & registro.item(0).selectSingleNode("./NomeCracha").text & "'"
				strSql = strSql & " , @Data_Nasc = '" & ano &"-"& dia & "-" & mes &"'"
				strSql = strSql & " , @Id_Edicao = '" & Session("cliente_edicao") & "'"
				strSql = strSql & " , @Id_Tipo_Credenciamento = '" & Session("cliente_tipo") & "'"
				strSql = strSql & " , @Id_Idioma = '" & Session("cliente_idioma") & "'"
			end if 
				'Exit For
			End If
			
		end if
		'Next
	'	response.write strSql
	'	response.end
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
	
	Function ConsultarCortesia(CPF)
		Dim xmlhttp
		Dim objXMLDoc
		Dim Url, Usuario, Senha, Metodo
		Dim strSql
		
		Url = "https://www.mbxeventos.com/wsAOL/Methods.asmx"
		Usuario = "ABF"
		Senha = "RmhUrD7E"
		Metodo = "EsCortesia"
		strSend = "<?xml version=""1.0"" encoding=""utf-8""?>" &_
			"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
			"<soap:Body><" & Metodo & " xmlns=""http://tempuri.org/""><user>abfne</user><pass>xAfr9ojG</pass><cpf>" & CPF & "</cpf></" & Metodo & "></soap:Body>" &_
			"</soap:Envelope>"
			
		strSend = "<?xml version=""1.0"" encoding=""utf-8""?><soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:tem=""http://tempuri.org/""> <soapenv:Header/> <soapenv:Body> <tem:" & Metodo & ">       <tem:user>" & Usuario & "</tem:user>   <tem:pass>" & Senha & "</tem:pass>   "

		
		
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
		if raiz.childNodes.item(0).childNodes.item(0).childNodes.item(0).text <> "" then
		
				ConsultarCortesia = raiz.childNodes.item(0).childNodes.item(0).childNodes.item(0).text
			
		End If
		
		Set objRS = Nothing
		Set NodeList = Nothing
		Set objXML = Nothing
		Set objXMLDoc = Nothing
	End Function
	
	Function SetComprador(CPF)
		Dim xmlhttp
		Dim objXMLDoc
		Dim Url, Usuario, Senha, Metodo
		Dim strSql
		
		Url = "https://www.mbxeventos.com/wsAOL/Methods.asmx"
		Usuario = "ABF"
		Senha = "RmhUrD7E"
		Metodo = "setComprador"
		strSend = "<?xml version=""1.0"" encoding=""utf-8""?>" &_
			"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
			"<soap:Body><" & Metodo & " xmlns=""http://tempuri.org/""><user>abfne</user><pass>xAfr9ojG</pass><cpf>" & CPF & "</cpf></" & Metodo & "></soap:Body>" &_
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
%>
