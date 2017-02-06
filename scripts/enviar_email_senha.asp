<%
Function Enviar_Email_Senha (id_edicao, id_idioma, CNPJ, Razao_Social, Email, Nome, Senha, Tipo, Pedido, Var_Visitante)

'	 // {1} = Email
'    // {2} = ID_Rel_Cadastro
'    // {3} = CPF
'    // {4} = Nome
'    // {5} = Cargo
'    // {6} = Depto
'    // {7} = CNPJ
'    // {8} = Razao

'id_edicao = 56

If Request("admin") = "sim" Then
	Response.Write(Tipo)
	'Response.End()
End If

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
			Idioma = "1"
			SgIdioma = "PTB"
	End Select

	Pagina_ID = 8

	SQL_Textos	=	" Select " &_
					"	ID_Texto, " &_
					"	ID_Tipo, " &_
					"	Identificacao, " &_
					"	Texto, " &_
					"	URL_Imagem " &_
					" From Paginas_Textos " &_
					" Where  " &_
					"	ID_Idioma = " & idioma & " " &_
					"	AND ID_Pagina = " & Pagina_ID & " " &_
					" Order By Ordem "
	'response.write(SQL_Textos)
	Set RS_Textos = Server.CreateObject("ADODB.Recordset")
	RS_Textos.Open SQL_Textos, Conexao

	If not RS_Textos.BOF or not RS_Textos.EOF Then
		total_registros = 0
		While not RS_Textos.EOF
			total_registros = total_registros + 1
			RS_Textos.MoveNext
		Wend
		RS_Textos.MoveFirst
		ReDim textos_array(total_registros-1)
		n = 0
		While not RS_Textos.EOF
			id = RS_Textos("id_texto")
			ident = RS_Textos("identificacao")
			texto = RS_Textos("texto")
			url_img = RS_Textos("url_imagem")
			textos_array(n) = Array(id, ident, texto, url_img)
			n = n + 1
			RS_Textos.MoveNext
		Wend
		RS_Textos.Close
	End If

	email_senha		= ""
	html		 	= ""
	assunto		 	= ""

	' Verifica Texto a ser apresentado ABF e ABF Nordeste
	Select Case (ID_Edicao)
		Case "22" ' ABF
			texto = textos_array(11)(2)
			texto_rodape = textos_array(12)(2)
		Case "35" ' ABF Nordeste
			texto = textos_array(13)(2)
			texto_rodape = textos_array(14)(2)
		Case Else ' Demais feiras
			texto = textos_array(1)(2)
			texto_rodape = textos_array(10)(2)
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
	'response.Write SQL_Edicoes_Configuracao
	'response.End
	Set RS_Edicoes_Configuracao = Server.CreateObject("ADODB.Recordset")
	RS_Edicoes_Configuracao.Open SQL_Edicoes_Configuracao, Conexao, 3, 3


	CNPJMask 		= Mid(CNPJ,1,2) & "." & Mid(CNPJ,4,3) & "." & Mid(CNPJ,7,3) & "/" & Mid(CNPJ,9,4) & "-" & Mid(CNPJ,13,2)

	If Tipo = "Recuperar_Senha" Then

		email_senha = 		"<html><head><meta http-equiv='Content-Type' content='text/html; charset=UTF-8' /><title>Credenciamento OnLine BTS</title></head>" &_
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
							"<tr>" &_
							"<td class='bemvindo'><div align='left'>" &_
							"<p align='center'><font size='2' face='verdana'><b>" & Nome & "</b>, segue seu C&oacute;digo de identifica&ccedil;&atilde;o solicitado para acesso ao sistema de compra de ingresso antecipada para <strong>" & RS_Edicoes_Configuracao("Feira") & " " & RS_Edicoes_Configuracao("Ano") & "</strong>.<br /></font></p>" &_
							"</div></td>" &_
							"</tr>" &_
							"<tr><td>&nbsp;</td></tr>" &_
							"<tr>" &_
							"<td>" &_
							"<table width='100%' border='0' cellspacing='2' cellpadding='2'>" &_
							"<tr>" &_
							"<td align='center'><font size='2' face='verdana' align='right'><strong>C&oacute;digo de identifica&ccedil;&atilde;o:&nbsp;" & senha & "</strong></font></td>" &_
							"</tr>" &_
							"<tr>" &_
							"<td>&nbsp;</td>" &_
							"</tr>" &_
							"</table>" &_
							"</body>" &_
							"</html>"

		html = email_senha
		assunto = "Recuperar Acesso - " & RS_Edicoes_Configuracao("Feira") & " " & RS_Edicoes_Configuracao("Ano") & " - Credenciamento OnLine BTS Informa"

	ElseIf Tipo = "Enviar_Ticket" Then

		SQL_Consulta_Pedidos =	"Select " &_
								"	P.*, " &_
								"	PH.* " &_
								"From Pedidos As P " &_
								"Inner Join Pedidos_Historico as PH " &_
								"	On P.Numero_Pedido = PH.Numero_Pedido " &_
								"Where " &_
								"	P.Numero_Pedido = '" & Pedido & "'  " &_
								"	And P.Status_Pedido = 3"
		'Response.Write(SQL_Consulta_Pedidos)
		'Response.End
		Set RS_Consulta_Pedidos = Server.CreateObject("ADODB.Recordset")
		RS_Consulta_Pedidos.Open SQL_Consulta_Pedidos, Conexao, 3, 3

		If Not RS_Consulta_Pedidos.Eof Then



			Tickets 			= True
			Numero_Pedido 		= RS_Consulta_Pedidos("Numero_Pedido")
			Numero_Transacao 	= RS_Consulta_Pedidos("Numero_Transacao")
			Cod_Autorizacao		= RS_Consulta_Pedidos("Codigo_Autorizacao")
			Visitante_ID		= RS_Consulta_Pedidos("ID_Visitante")
			ID_Pedido			= RS_Consulta_Pedidos("ID_Pedido")
			Vlr_Pedido 			= RS_Consulta_Pedidos("Valor_Pedido")

			'Response.Write Visitante_ID
			'response.End
			If Cstr(Visitante_ID) <> Cstr(Var_Visitante) Then
				Valor_Pedido		= FormatNumber(Vlr_Pedido,2)
			Else
				Valor_Pedido		= FormatNumber(Vlr_Pedido,2)
			End If

		End If

'response.Write RS_Edicoes_Configuracao("Logo_Box")
'response.End
		comprovante = 	"<html><head><meta http-equiv='Content-Type' content='text/html; charset=ISO-8859-1' /><title>Credenciamento OnLine BTS</title></head>" &_
						"<body>" &_
						"<table width='520' border='0' align='center' cellpadding='0' cellspacing='0'>" &_
						"<tr>" &_
						"<td><img src='http://credenciamento.btsinforma.com.br/images/informa_exhibition.png' alt='' width='95' height='52' hspace='15' /></td>" &_
						"<td width='15'>&nbsp;</td>" &_
						"<td align='right'><img src='http://credenciamento.btsinforma.com.br" & RS_Edicoes_Configuracao("Logo_Box") & "'/></td>" &_
						"</tr>" &_
						"</table>" &_
						"<br /><div style=""font-size:17px; color:#1F497D; font-weight:bold; font-family:'Calibri','sans-serif'; text-align:center;"">Sua compra foi realizada com sucesso! Retire seu ingresso nos guichês de atendimento na entrada da ABF Franchising Expo 2016 – Expo Center Norte</div><br /><br />" &_
						"<table width='520' border='0' align='center' cellpadding='0' cellspacing='0'> " &_
						"<tr><td class='bemvindo'>&nbsp;</td></tr>" &_
						"<tr>" &_
						"<td class='bemvindo'>" &_
						"<table width='100%' border='0' cellspacing='2' cellpadding='2'>" &_
						"<tr>" &_
						"<td width='220'><font size='2' face='verdana'><strong>Pagamento:</strong></font></td>" &_
						"<td ><font size='3' face='verdana'><strong>Aprovado</div></strong></font></td>" &_
						"</tr>" &_
						"<tr>" &_
						"<td width='220'><font size='2' face='verdana'><strong>Numero do Pedido:</strong></font></td>" &_
						"<td ><font size='3' face='verdana'><strong>" & Numero_Pedido & "</div></strong></font></td>" &_
						"</tr>" &_
						"<tr> " &_
						"<td width='220'><font size='2' face='verdana'><strong>Transa&ccedil;&atilde;o:</strong></font></td> " &_
						"<td ><font size='3' face='verdana'><strong>" & Numero_Transacao & "</div></strong></font></td> " &_
						"</tr> " &_
						"<tr> " &_
						"<td width='220'><font size='2' face='verdana'><strong>C&oacute;d. da Autoriza&ccedil;&atilde;o:</strong></font></td> " &_
						"<td ><font size='3' face='verdana'><strong>" & Cod_Autorizacao & "</div></strong></font></td> " &_
						"</tr> " &_
						"<tr> " &_
						"<td width='220'><font size='2' face='verdana'><strong>Valor Pago:</strong></font></td> " &_
						"<td><font size='3' face='verdana'><strong>R$ " & Valor_Pedido & "</div></strong></font></td> " &_
						"</tr> " &_
						"</table> <br><br>" &_
						"<table width='520' border='0' align='center' cellpadding='0' cellspacing='0'> " &_
						"<tr> " &_
						"<td style=' border-bottom: 1px dotted #ccc'><b><font size='2' face='verdana'>NOME COMPLETO</font></b></td>" &_
						"<td style=' border-bottom: 1px dotted #ccc'><b><font size='2' face='verdana'>TIPO</font></b></td>" &_
						"<td style=' border-bottom: 1px dotted #ccc'><b><font size='2' face='verdana'>DOCUMENTO</font></b></td>" &_
						"</tr>"

						'response.Write Var_Visitante
						'response.End
		' Nao realizei nenhum PEDIDO
		If Cstr(Visitante_ID) <> Cstr(Var_Visitante) Then

			SQL_Carrinho = 	"Select " &_
							"	C.ID_Carrinho,  " &_
							"	C.ID_Visitante,  " &_
							"	C.ID_Pedido,  " &_
							"	C.ID_Rel_Cadastro, " &_
							"	P.Status_Pedido, " &_
							"	V.Nome_Completo, " &_
							"	V.CPF, " &_
							"	V.Passaporte, " &_
							"	V.Email " &_
							"From  Pedidos_Carrinho  As C " &_
							"Inner Join Visitantes As V On V.ID_Visitante = C.ID_Visitante " &_
							"Inner Join Pedidos As P On P.ID_Pedido = C.ID_Pedido " &_
							"Where " &_
							"	C.ID_Pedido = " & ID_Pedido & " " &_
							"	And C.ID_Visitante = " & Var_Visitante & ""


		' Realizei este pedido
		Else
			SQL_Carrinho = 	"Select " &_
							"	C.ID_Carrinho,  " &_
							"	C.ID_Visitante,  " &_
							"	C.ID_Pedido,  " &_
							"	C.ID_Rel_Cadastro, " &_
							"	P.Status_Pedido, " &_
							"	V.Nome_Completo, " &_
							"	V.CPF, " &_
							"	V.Passaporte, " &_
							"	V.Email " &_
							"From  Pedidos_Carrinho  As C " &_
							"Inner Join Visitantes As V On V.ID_Visitante = C.ID_Visitante " &_
							"Inner Join Pedidos As P On P.ID_Pedido = C.ID_Pedido " &_
							"Where " &_
							"	C.ID_Pedido = " & ID_Pedido & " " &_
							"	And C.Cancelado = 0"

		End If

		'Response.Write(SQL_Carrinho)
		'REsponse.End()

		Set RS_Carrinho = Server.CreateObject("ADODB.Recordset")
		RS_Carrinho.Open SQL_Carrinho, Conexao, 3, 3

		Primeiro = 0
		Z = True

		If Not RS_Carrinho.Eof Then
			While Not RS_Carrinho.Eof

			Email_Carrinho = RS_Carrinho("Email")

			If Len(Trim(RS_Carrinho("CPF"))) > 0 Then
            	Tipo_Documento = "CPF"
            Else
            	Tipo_Documento = "Passaporte"
            End If

			If Len(Trim(RS_Carrinho("CPF"))) > 0 Then
				Numero_Documento = RS_Carrinho("CPF")
			Else
				Numero_Documento = RS_Carrinho("Passaporte")
			End If

		comprovante = comprovante & "<tr bgcolor='" & Cor_Fundo & "' style='padding: 5px; font-weight: 100'>"
		comprovante = comprovante & "<td style='padding: 5px; width: 375px; border-bottom: 1px dotted #ccc'><font size='2' face='verdana'>" & RS_Carrinho("Nome_Completo") & "</font></td>"
		comprovante = comprovante & "<td style='padding: 5px; width: 100px; border-bottom: 1px dotted #ccc'><font size='2' face='verdana'>" & Tipo_Documento & "</font></td>"
		comprovante = comprovante & "<td style='padding: 5px; width: 100px; border-bottom: 1px dotted #ccc'><font size='2' face='verdana'>" & Numero_Documento & "</font></td>"
		comprovante = comprovante & "</tr>"

			RS_Carrinho.MoveNext
			Wend
		End If

		comprovante = comprovante & "</table>"
		comprovante = comprovante & "<br/>"
		comprovante = comprovante & "<table width='520' border='0' align='center' cellpadding='0' cellspacing='0'><tr><td><font size='1' face='verdana'>"
		comprovante = comprovante & "<div align='center'><img src='http://credenciamento.btsinforma.com.br/images/8480_1288x241_port.jpg'/></div><br><br>"
		comprovante = comprovante & "- Para retirar seu ingresso e a credencial para acesso ao evento, tenha em mãos seu comprovante de compra e seu CPF.<br>"
		comprovante = comprovante & "- O ingresso é pessoal e intransferível, sendo obrigatória a apresentação do CPF para sua retirada.<br>"
		comprovante = comprovante & "- Não será permitida a entrada de pessoas trajando bermudas, camiseta regata e/ou chinelos.<br>"
		comprovante = comprovante & "- Proibida a entrada de menores de 16 anos desacompanhados.<br>"
		comprovante = comprovante & "</font></td></tr></table>"
		comprovante = comprovante & "</body>"
		comprovante = comprovante & "</html>"


		html = comprovante
		assunto = "Comprovante de Pagamento - " & RS_Edicoes_Configuracao("Feira") & " " & RS_Edicoes_Configuracao("Ano") & " - Pedido: " & Pedido
	'Response.Write assunto
	'response.end
	End If

	RS_Edicoes_Configuracao.Close

	Set Mail = Server.CreateObject("Persits.MailSender")

	Mail.CharSet = "ISO-8859-1"

		Mail.CharSet = "ISO-8859-1"
		Mail.Username = "btsinforma"
		Mail.Password = "Phebru28"
		Mail.host = "smtpcorp.com"
		Mail.port = 2525

	Mail.From = "brazilexhibitorsmanual@informa.com" ' Required

	Mail.FromName = "Brazil Exhibitors Manual" ' Optional

	Mail.AddAddress trim(Email)
	'Mail.AddBCC Trim(RS_Verifica("email_copia"))
	'Mail.AddBCC "guilherme.ribeiro@informa.com" 				' cópia TI
	Mail.AddBCC "gabriel.petro@informa.com" 				' cópia ATENDIMENTO
	Mail.AddBCC "andre.alves@informa.com" 				' cópia TI

	Mail.Subject = assunto
	Mail.Body = html
	Mail.IsHTML = True


	Mail.Send

	'// Alexandre Fischer - 17/04/2014 - Utilizar CDO.SYS para envio de e-mails autenticados
	'Dim objCDOSYSMail
	'Dim objCDOSYSCon
    '
	'Set objCDOSYSMail = Server.CreateObject("CDO.Message")
	'Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")
    '
	'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtpcorp.com"
	'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 2525
	'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
	'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "btsinforma"
	'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Phebru28"
	'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
    '
	'objCDOSYSCon.Fields.Update
	'
	'If Email = "" Then
	'	Email = Email_Carrinho
	'End If
    '
	'Set objCDOSYSMail.Configuration = objCDOSYSCon
	'objCDOSYSMail.From = """Credenciamento OnLine BTS Informa""<credenciamento@btsinforma.com.br>"
	'objCDOSYSMail.To = trim(Email)
	'objCDOSYSMail.Bcc = "gabriel.petro@informa.com"
	'objCDOSYSMail.Subject = assunto
	'objCDOSYSMail.HtmlBody = html
    '
	'objCDOSYSMail.Send
	'response.write Email
	'response.end
'							response.write id_edicao
'response.end

End Function
%>
