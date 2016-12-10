<%
	Function ConsultarWS(CPF)
		Dim xmlhttp
		Dim objXMLDoc
		Dim Url, Usuario, Senha, Metodo
		Dim strSql
		
		Url = "https://www.mbxeventos.com/wsAOL/Methods.asmx"
		'Url = "https://www.mbxeventos.com/wsConvidados/Methods.asmx
		Usuario = "ABF"
		'Usuario = "BTS"
		Senha = "RmhUrD7E"
		Metodo = "getXMLCPF"
		strSend = "<?xml version=""1.0"" encoding=""utf-8""?>" &_
			"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
			"<soap:Body><" & Metodo & " xmlns=""http://tempuri.org/""><user>abfne</user><pass>xAfr9ojG</pass><cpf>" & CPF & "</cpf></" & Metodo & "></soap:Body>" &_
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
		For Each Node In NodeList
			If Node.childNodes(1).Text = CPF Then
				strSql = "SET DATEFORMAT YDM; EXEC dbo.SP_IN_VISITANTES_CADASTRO"
				strSql = strSql & " @CPF = '" & Node.childNodes(1).Text & "'"
				strSql = strSql & " , @Email = '" & Node.childNodes(2).Text & "'"
				strSql = strSql & " , @Nome_Completo = '" & Node.childNodes(3).Text & "'"
				strSql = strSql & " , @Nome_Credencial = '" & Node.childNodes(4).Text & "'"
				strSql = strSql & " , @Data_Nasc = '" & Left(Node.childNodes(14).Text, 10) & "'"
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
%>
