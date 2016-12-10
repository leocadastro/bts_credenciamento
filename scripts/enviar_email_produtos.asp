<%
Function Enviar_Email_Produtos (id_edicao, id_idioma, CNPJ, Razao_Social, Email, Nome, ID_Empresa, DDI, DDD, Telefone)
'	 // {1} = Email
'    // {2} = ID_Rel_Cadastro
'    // {3} = CPF
'    // {4} = Nome
'    // {5} = Cargo
'    // {6} = Depto
'    // {7} = CNPJ
'    // {8} = Razao

Idioma 	= id_idioma

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
								"	EC.ID_Edicao = " & id_edicao & " " &_
								"	AND EC.Ativo = 1"
	'response.write(SQL_Edicoes_Configuracao)
	Set RS_Edicoes_Configuracao = Server.CreateObject("ADODB.Recordset")
	RS_Edicoes_Configuracao.CursorType = 0
	RS_Edicoes_Configuracao.LockType = 1
	RS_Edicoes_Configuracao.Open SQL_Edicoes_Configuracao, Conexao									


	CNPJMask 		= Mid(CNPJ,1,2) & "." & Mid(CNPJ,4,3) & "." & Mid(CNPJ,7,3) & "/" & Mid(CNPJ,9,4) & "-" & Mid(CNPJ,13,2)

	SQL_Produtos = 		"Select " &_
						"	Principal_Produto as Produto  " &_
						"FROM  " &_
						"	Relacionamento_Produto_Edicao_Empresa_Visitante_v2  " &_
						"Where  " &_
						"	ID_Empresa =  " & ID_Empresa & " " &_
						"Order by Principal_Produto "
	Set RS_Produtos = Server.CreateObject("ADODB.Recordset")
	RS_Produtos.CursorType = 0
	RS_Produtos.LockType = 1
	RS_Produtos.Open SQL_Produtos, Conexao	

	lista_produtos = ""
	If not RS_Produtos.BOF or not RS_Produtos.EOF Then
		lista_produtos = "<ul style='list-style:decimal-leading-zero'>"
		While not RS_Produtos.EOF
			lista_produtos = lista_produtos & "<li>"
			lista_produtos = lista_produtos & Trim(RS_Produtos("Produto"))
			lista_produtos = lista_produtos & "</li>"
			RS_Produtos.MoveNext
		Wend
		RS_Produtos.Close
		lista_produtos = lista_produtos & "</ul>"
	End If

	email_produto		= ""
	html		 		= ""
	assunto		 		= ""
	
	
	email_produto =		"<html><head><meta http-equiv='Content-Type' content='text/html; charset=UTF-8' /><title>Credenciamento OnLine BTS</title></head>" &_
						"<body>" &_
						"<table width='520' border='0' align='center' cellpadding='0' cellspacing='0'>" &_
						"<tr>" &_
						"<td><img src='http://credenciamento.btsinforma.com.br/images/informa_exhibition.png' alt='' width='95' height='52' hspace='15' /></td>" &_
						"<td width='15'>&nbsp;</td>" &_
						"<td align='right'><img src='http://credenciamento.btsinforma.com.br" & RS_Edicoes_Configuracao("Logo_Box") & "'/></td>" &_
						"</tr>" &_
						"</table>" &_
						"<table width='520' border='0' align='center' cellpadding='0' cellspacing='0'> " &_
						"<tr><td class='bemvindo'>&nbsp;</td></tr>" &_
						"<tr><td class='bemvindo'>" &_
							"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='2' >" &_
							"<tr>" &_
								"<td><font size='2' face='verdana'><strong>CNPJ</strong></font></td>" &_
								"<td><font size='2' face='verdana'>" & CNPJMask & "</font></td>" &_
							"</tr>" &_
							"<tr>" &_
								"<td><font size='2' face='verdana'><strong>Empresa</strong></font></td>" &_
								"<td><font size='2' face='verdana'>" & Razao & "</font></td>" &_
							"</tr>" &_
							"<tr>" &_
								"<td><font size='2' face='verdana'><strong>Solicitante</strong></font></td>" &_
								"<td><font size='2' face='verdana'>" & Nome & "</font></td>" &_
							"</tr>" &_
							"<tr>" &_
								"<td><font size='2' face='verdana'><strong>E-mail</strong></font></td>" &_
								"<td><font size='2' face='verdana'>" & Email & "</font></td>" &_
							"</tr>" &_
							"<tr>" &_
								"<td><font size='2' face='verdana'><strong>Telefone</strong></font></td>" &_
								"<td><font size='2' face='verdana'>" & DDI & " (" & DDD & ") " & Telefone & "</font></td>" &_
							"</tr>" &_
							"</table><br>" &_
						"</td></tr>" &_
						"<tr><td class='bemvindo'>&nbsp;</td></tr>" &_
						"<tr>" &_
						"<td class='bemvindo'><div align='left'>" &_
						"<p align='left'><font size='2' face='verdana'><b>" & Nome & "</b>, recebemos sua solicitação para alterar os produtos de sua empresa, abaixo a lista dos produtos cadastrados até então:<br>" & lista_produtos & "<br /></font></p>" &_
						"</div></td>" &_
						"</tr>" &_
						"<tr><td>&nbsp;</td></tr>" &_
						"<tr>" &_
						"<td>" &_
						"<p align='center'><font size='2' face='verdana'>Em breve entraremos em contato.</font></p>" &_
						"</td>" &_
						"</tr>" &_
						"</table>" &_
						"</body>" &_
						"</html>"

'	Select Case (Idioma)
'		Case "1"
'			email_confirmacao = email_confirmacao_PTB
'		Case "2"
'			email_confirmacao = email_confirmacao_ENG
'		Case "3"
'			email_confirmacao = email_confirmacao_ESP
'		Case Else
'			email_confirmacao = email_confirmacao_PTB
'	End Select


			html = email_produto
			assunto = "Alterar Produtos - " & RS_Edicoes_Configuracao("Feira") & " " & RS_Edicoes_Configuracao("Ano") & " - Credenciamento OnLine BTS Informa"

	'		response.write(html)
	
	RS_Edicoes_Configuracao.Close
	
		'// Alexandre Fischer - 17/04/2014 - Utilizar CDO.SYS para envio de e-mails autenticados
	Dim objCDOSYSMail
	Dim objCDOSYSCon

	Set objCDOSYSMail = Server.CreateObject("CDO.Message")
	Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")
	
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtpcorp.com"
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 2525
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "btsinforma"
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Phebru28" 
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	
	objCDOSYSCon.Fields.Update 
	
	Set objCDOSYSMail.Configuration = objCDOSYSCon
	objCDOSYSMail.From = """Credenciamento OnLine BTS Informa""<credenciamento@informaexhibitionsbrasil.com.br>"
	objCDOSYSMail.To = Email
	objCDOSYSMail.Bcc = "andre.alves@informa.com"
	objCDOSYSMail.Subject = assunto
	objCDOSYSMail.HtmlBody = html
	
	On Error Resume Next
	
	objCDOSYSMail.Send
	
	Set objCDOSYSMail = Nothing
	Set objCDOSYSCon = Nothing
	
	If Err <> 0 Then
		'Response.Write "<br><div { 'Ocorreu um erro ao enviar o email: " & Err.Description & "' }"
		%>
        <br><div style="background-color:#000; color:#FFF;">{ Ocorreu um erro ao enviar o email: "<%=Err.Description%>" }</div>
        <%
	Else
		'Response.Write "<br>{ " & observacao & " - Email enviado para: " & email & " }"
		%>
        <br><div style="background-color:#000; color:#FFF;">{ obs: <%=observacao%> - Email <b>Produtos</b> enviado para: <%=email%> }</div>
        <%
	End If 
	
	On Error GoTo 0
	
End Function 
%>