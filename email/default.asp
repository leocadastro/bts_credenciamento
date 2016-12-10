<%

		email_boleto = 	"<html><head><meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'><title>BTS</title> " &_
						"</head><body><br /><table width='600' border='0' align='center' cellpadding='0' cellspacing='0' bgcolor='#FFFFFF'> " &_
						"<tr><td colspan='2' align='center'>&nbsp;</td></tr><tr><td width='250' height='80' align='center'> " &_
						"<a href='http://www.manualexpositor.com.br/'><img src='http://manual.btsinforma.com.br/cabecalho_logo_btsinforma.gif' width='188'  border='0' /></a> " &_
						"</td><td height='80' align='center' valign='bottom'><div style='font-size:20px;font-family:Tahoma, Verdana, Segoe, sans-serif'><B>EXHIBITOR MANUAL ONLINE</B></div></td></tr></table> " &_
						"<table width='600' border='0' align='center' cellpadding='0' cellspacing='0'><tr><td height='60' bgcolor='#FFFFFF'>&nbsp;</td> " &_
						"</tr><tr><td align='center' bgcolor='#FFFFFF'><table width='570' border='0' align='center' cellpadding='0' cellspacing='0'> " &_
						"<tr><td valign='top' style='font-family:Verdana, Geneva, sans-serif; font-size:13px; color:#000;'> " &_
						"Dear Exhibitor, <br><BR><b> Nome </b><br /><BR>" &_
						"Access to the <strong>Online Exhibitor Manual BTS Informa</strong> is now available for filling in the forms of <strong>Feira Feira.</strong><br /><br />" &_
						"<b>Access</b> <a href= 'http://www.manualexpositor.com.br/' >http://www.manualexpositor.com.br/</a> <br /><br />" &_
						"<strong>Contract Number:</strong> Codigo<br> <br />" &_
						"Stay tuned to all the rules and deadlines for contracting services and submission of the requested documents.<br /><br />" &_
						"In case of doubt, the Service Center will be happy to serve you.<br /><br />" &_
						"Contact: +55 (11) 000000000 <br /><br />" &_
						"E-mail: <a href= 'mailto:andresistinfo@gmail' >email</a><br /><br /><br />" &_
						"<span style='font-size:10px; text-align:center; color:#F00;'> Note: This email has been automatically,  " &_
						"no response is required.</span> " &_
						"</td></tr></table></td></tr><tr><td align='center' bgcolor='#FFFFFF'>&nbsp;</td></tr></table></body></html>" 
						
						assunto = " Informa | System Access - Feira"
'response.Write email_boleto
'response.end

			html = email_boleto

	Set Mail = Server.CreateObject("Persits.MailSender")
	
	
	
	
	Mail.CharSet = "ISO-8859-1"
	
	'Mail.Username = "workflow@informagroup.com.br"
	'Mail.Password = "drUmafe7"
	'Mail.host = "smtp.informagroup.com.br"
	'Mail.port = 587
	'Mail.From = "workflow@informagroup.com.br"
	
		'Mail.CharSet = "ISO-8859-1"
		Mail.Username = "btsinforma"
		Mail.Password = "Phebru28"
		Mail.host = "smtpcorp.com"
		Mail.port = 2525
'response.Write Mail.CharSet
	'response.end
	Mail.From = "BrazilExhibitorsManual@informa.com" ' Required
	'Mail.To
	Mail.FromName = "Teste Servidor 14" ' Optional 
	Mail.AddAddress("andresistinfo@gmail.com")
	Mail.AddBCC ("andre.alves@informa.com") 				' cópia TI
	
	Mail.Subject = assunto
	Mail.Body = html
	Mail.IsHTML = True 

	'response.Write Mail.From
	'response.end
		Mail.Send

		
%>