<%

	Set Conexao_TC = Server.CreateObject("ADODB.Connection")
	Conexao_TC.Open Application("cnn")


	SQL_Valor_Pedidos = "select codigo, nome, ContatoEmail, enviado, feira, telefone, email_feira, email_copia from envio_email where enviado = 0"
	'Response.Write(Conexao_TC.ConnectionString)
	'Response.End
	Set RS_Valor_Pedidos = Server.CreateObject("ADODB.Recordset")
	

RS_Valor_Pedidos.Open SQL_Valor_Pedidos, "Provider=SQLOLEDB;Password=Bem123;Persist Security Info=True;User ID=Bem;Initial Catalog=ExhibitorsManual;Data Source=172.21.43.15, 63946", 1, 3
	'RS_Valor_Pedidos.Open SQL_Valor_Pedidos, "Provider=SQLNCLI11.1;Integrated Security="""";Persist Security Info=true;User ID=Bem;PWD=Bem123;Initial Catalog=ExhibitorsManual;Data Source=172.21.43.15, 63946;Initial File Name="""";Server SPN=""""", 3, 3

do while not RS_Valor_Pedidos.EOF	 

		email_boleto = 	"<html><head><meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'><title>BTS</title> " &_
						"</head><body><br /><table width='600' border='0' align='center' cellpadding='0' cellspacing='0' bgcolor='#FFFFFF'> " &_
						"<tr><td colspan='2' align='center'>&nbsp;</td></tr><tr><td width='250' height='80' align='center'> " &_
						"<a href='http://manualexpositor.btsinforma.com.br/'><img src='http://cs.btsinforma.com.br/img/geral/cabecalho_logo_btsinforma.gif' width='188' height='103' border='0' /></a> " &_
						"</td><td height='80' align='center' valign='bottom'><div style='font-size:20px;font-family:Tahoma, Verdana, Segoe, sans-serif'><B>MANUAL DO EXPOSITOR ONLINE</B></div></td></tr></table> " &_
						"<table width='600' border='0' align='center' cellpadding='0' cellspacing='0'><tr><td height='60' bgcolor='#FFFFFF'>&nbsp;</td> " &_
						"</tr><tr><td align='center' bgcolor='#FFFFFF'><table width='570' border='0' align='center' cellpadding='0' cellspacing='0'> " &_
						"<tr><td valign='top' style='font-family:Verdana, Geneva, sans-serif; font-size:13px; color:#000;'> " &_
						"Prezado Expositor, <br><BR><b> " & RS_Valor_Pedidos("nome") & "</b><br /><BR>" &_
						"O acesso ao <strong>Manual do Expositor Online</strong> da BTS Informa j&aacute; est&aacute; dispon&iacute;vel para preenchimento dos formul&aacute;rios da <strong>Feira " & RS_Valor_Pedidos("feira") & ".</strong><br /><br />" &_
						"<b>Acesse</b> <a href= 'http://manualexpositor.btsinforma.com.br/' >http://manualexpositor.btsinforma.com.br/</a> <br /><br />" &_
						"<strong>N&uacute;mero de Contrato:</strong> " & RS_Valor_Pedidos("codigo") & "<br> <br />" &_
						"Fique atento a todas as normas e prazos para contrata&ccedil;&atilde;o de servi&ccedil;os e envio dos documentos solicitados.<br /><br />" &_
						"Em caso de d&uacute;vidas, a Central de Atendimento estar&aacute; &agrave; disposi&ccedil;&atilde;o para atend&ecirc;-lo.<br /><br />" &_
						"Contato: (11) " & RS_Valor_Pedidos("telefone") & " <br /><br />" &_
						"E-mail: <a href= 'mailto:" & RS_Valor_Pedidos("email_feira") & "' >" & RS_Valor_Pedidos("email_feira") & "</a><br /><br /><br />" &_
						"<span style='font-size:10px; text-align:center; color:#F00;'> Aten&ccedil;&atilde;o: esse e-mail foi gerado automaticamente,  " &_
						"n&atilde;o &eacute; necess&aacute;ria nenhuma resposta.</span> " &_
						"</td></tr></table></td></tr><tr><td align='center' bgcolor='#FFFFFF'>&nbsp;</td></tr></table></body></html>" 

			html = email_boleto
			assunto = " Informa | Acesso ao Sistema - " & RS_Valor_Pedidos("feira")
	
	Set Mail = Server.CreateObject("Persits.MailSender")
	
	Mail.CharSet = "ISO-8859-1"
	
	Mail.Username = "btsinforma"
	Mail.Password = "Phebru28"
	Mail.host = "smtpcorp.com"
	Mail.port = 2525

	Mail.From =  RS_Valor_Pedidos("email_feira") ' Required
	
	Mail.FromName = "Brazil Exhibitors Manual" ' Optional 
	
	'Mail.AddAddress  "thiago.souza@informa.com"
	Mail.AddAddress RS_Valor_Pedidos("ContatoEmail")
	Mail.AddBCC "gabriel.petro@informa.com" 				' cópia TI
	Mail.AddBCC "andre.alves@informa.com" 				' cópia TI
	
	If RS_Valor_Pedidos("email_copia") <> "" or IsNull(RS_Valor_Pedidos("email_copia")) then
		Mail.AddBCC RS_Valor_Pedidos("email_copia")			' cópia ATENDIMENTO
	end if
	
	Mail.Subject = assunto
	Mail.Body = html
	Mail.IsHTML = True 

	
		Mail.Send
		
		RS_Valor_Pedidos("enviado") = 1
		RS_Valor_Pedidos.update
	'End If
	
	RS_Valor_Pedidos.MOVENEXT
	LOOP
	
	'On Error Resume Next
	
	



%>