<%
Function Enviar_Email (id_edicao, id_idioma, ID_Formulario, Email, ID_Rel_Cadastro, CPF, Passaporte, Nome, Cargo, Depto, CNPJ, Razao)
'	 // {1} = Email
'    // {2} = ID_Rel_Cadastro
'    // {3} = CPF
'    // {4} = Nome
'    // {5} = Cargo
'    // {6} = Depto
'    // {7} = CNPJ
'    // {8} = Razao

Idioma 						= id_idioma
Formulario_Credenciamento 	= ID_Formulario

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


'	For i = Lbound(textos_array) to Ubound(textos_array)
'		response.write("[ i: " & i & " ] [ ident: " & textos_array(i)(1) & " ]  [ txt: " & textos_array(i)(2) & " ]  [ img: " & textos_array(i)(3) & " ]<br>")
'	Next
'===========================================================
%>
<% If Request("teste") = "s" Then %>
	<!--#include virtual="/includes/exibir_array.asp"-->
<% End IF

	' Buscar nome do FORM
	SQL_Nome_Formulario = 	"Select " &_
							"	ID_Formulario " &_
							"	,Nome " &_
							"FROM Formularios " &_
							"Where ID_Formulario = " & Formulario_Credenciamento
	Set RS_Nome_Formulario = Server.CreateObject("ADODB.Recordset")
	RS_Nome_Formulario.CursorType = 0
	RS_Nome_Formulario.LockType = 1
	RS_Nome_Formulario.Open SQL_Nome_Formulario, Conexao

	Nome_Formulario = ""
	If not RS_Nome_Formulario.BOF or not RS_Nome_Formulario.EOF Then
		Nome_Formulario = RS_Nome_Formulario("nome")
	End If

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

	ID_Cadastro		= limpar_texto(ID_Rel_Cadastro)
	For i = Len(ID_Cadastro)+1 To 6
		ID_Cadastro = "0" & ID_Cadastro
	Next

	' Correcao de informacao aprensetada no campo
	CPFEmail	= limpar_texto(Request("frmCPF"))
	CPFEmail	= Replace(CPFEmail,".","")
	CPFEmail	= Replace(CPFEmail,"-","")

	if len(Trim(CPFEmail)) > 0 Then
		CPFMask			= Mid(CPFEmail,1,3) & "." & Mid(CPFEmail,4,3) & "." & Mid(CPFEmail,7,3) & "-" & Mid(CPFEmail,10,2)
	Else
		CPFMask			= limpar_texto(Request("frmPassaporte"))
	End If

	if ID_Formulario <> 4 Then
		' Select de Cargos
		ID_Cargo	 		= limpar_texto(Request("frmCargo"))
		SQL_Cargo 		= "SELECT " &_
							"	ID_Cargo as Id, " &_
							"	Cargo_" & SgIdioma & " as Cargo " &_
							"FROM Cargo " &_
							"WHERE " &_
							"	Ativo = 1 " &_
							"	AND ID_Cargo = " & ID_Cargo & " "
		Set RS_Cargo = Server.CreateObject("ADODB.Recordset")
		RS_Cargo.CursorType = 0
		RS_Cargo.LockType = 1
		RS_Cargo.Open SQL_Cargo, Conexao
		
		Email_Cargo = RS_Cargo("Cargo")

		' Select de Departamentos
		ID_Depto = limpar_texto(Request("frmDepto"))
		SQL_Depto 		= "SELECT " &_
							"	ID_Depto as Id, " &_
							"	Depto_" & SgIdioma & " as Depto " &_
							"FROM Depto " &_
							"WHERE " &_
							"	Ativo = 1 " &_
							"	AND ID_Depto = " & ID_Depto & "  "
		Set RS_Depto = Server.CreateObject("ADODB.Recordset")
		RS_Depto.CursorType = 0
		RS_Depto.LockType = 1
		RS_Depto.Open SQL_Depto, Conexao

		Email_Depto = RS_Depto("Depto")

	End If

	CNPJMask 		= Mid(CNPJ,1,2) & "." & Mid(CNPJ,3,3) & "." & Mid(CNPJ,6,3) & "/" & Mid(CNPJ,9,4) & "-" & Mid(CNPJ,13,2)

	hoje 	 = day(Now) & " / " & Ucase(Left(MonthName(Month(Now)),3)) & " / " & Year(Now)
	hoje_eng = Year(Now) & " / " & Ucase(Left(MonthName(Month(Now)),3)) & " / " & day(Now)
	
	email_confirmacao	= ""
	html		 		= ""
	assunto		 		= ""

    ' Alerta NORDESTE
    Select Case (ID_Edicao)
        Case "34"
            TextoAlerta     =   "• Data para visita de Instituições de Ensino: 08/11/2013.<br>"
        Case "33"
            TextoAlerta     =   "• Data para visita de Instituições de Ensino: 08/11/2013<br>."
        Case Else
            TextoAlerta     =   ""
    End Select

	' Verifica Texto a ser apresentado ABF e ABF Nordeste
	Select Case (ID_Edicao)
		Case "46" ' ABF
			texto = textos_array(11)(2)
			texto_rodape = textos_array(12)(2)
		Case "32" ' ABF Nordeste
			texto = textos_array(18)(2)
			texto_rodape = textos_array(19)(2)
		Case "35" ' ABF NE 2013
			texto = textos_array(20)(2)
			texto_rodape = textos_array(10)(2)
		Case "45" ' FORMOBILE 2014
			texto = textos_array(21)(2)
			texto_rodape = textos_array(22)(2)	
		Case Else ' Demais feiras
			texto = textos_array(1)(2)
			texto_rodape = textos_array(10)(2)
	End Select

	' Nova Implementacao - 07/05/2014 - Leandro Santiago | HD 8234
    ' INICIO - Texto especifico para as feiras

	' Fispa Tecnologia Universidade 2014
	if ID_Edicao = "47"and Idioma = "1" and ID_Formulario = 5 Then
		texto = textos_array(24)(2)
		texto_rodape = textos_array(25)(2)
	End If

	' Fispa Service Universidade 2014
	if ID_Edicao = "48" and Idioma = "1" and ID_Formulario = 5 Then
		texto = textos_array(26)(2)
		texto_rodape = textos_array(27)(2)
	End If

	' Fispa Cafe Universidade 2014
	if ID_Edicao = "49" and Idioma = "1" and ID_Formulario = 5 Then
		texto = textos_array(26)(2)
		texto_rodape = textos_array(27)(2)
	End If

	' Fispa Sorvete Universidade 2014
	if ID_Edicao = "50" and Idioma = "1" and ID_Formulario = 5 Then
		texto = textos_array(26)(2)
		texto_rodape = textos_array(27)(2)
	End If

	' INICIO - Texto especifico para as feiras

	email_confirmacao = "<html><head><meta http-equiv='Content-Type' content='text/html; charset=ISO-8859-1' /><title>Credenciamento OnLine BTS</title></head>" &_
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
						"<p align='center'><font size='2' face='verdana'>" & textos_array(0)(2) & "<br />" &_
						"" & texto & "<br />" &_
						"" & textos_array(2)(2) & " "
	If ID_Edicao = "1" or ID_Edicao = "10" or ID_Edicao = "2" Then
	email_confirmacao = email_confirmacao &_ 
						"<br/><br/>Garanta as melhores condições em sua viagem e hospedagem <a href='http://www.almax.com.br/'' target='blank'>clique aqui</a>"
	End If
	email_confirmacao = email_confirmacao &_ 
						"</font></p>" &_
						"</div></td>" &_
						"</tr>" &_
						"<tr><td>" & TextoAlerta & "</td></tr>" &_
						"<tr>" &_
						"<td>" &_
						"<table width='100%' border='0' cellspacing='2' cellpadding='2'>" &_
						"<tr>" &_
						"<td width='220'><font size='2' face='verdana'><strong>" & textos_array(3)(2) & "</strong></font></td>" &_
						"<td ><font size='3' face='verdana'><strong>" & ID_Cadastro & "</strong></font></td>" &_
						"</tr>" &_
						"<tr>" &_
						"<td><font size='2' face='verdana'><strong>" & textos_array(4)(2) & "</strong></font></td>" &_
						"<td><font size='2' face='verdana'>" & CPFMask & "</td>" &_
						"</tr>" &_
						"<tr>" &_
						"<td><font size='2' face='verdana'><strong>" & textos_array(5)(2) & "</strong></font></td>" &_
						"<td><font size='2' face='verdana'>" & Nome & "</font></td>" &_
						"</tr>" &_
						"<tr>"
	If Idioma = "1" Then
		If ID_Formulario <> 4 Then 
			email_confirmacao = email_confirmacao &_
							"<td><font size='2' face='verdana'><strong>" & textos_array(6)(2) & "</strong></font></td>" &_
							"<td><font size='2' face='verdana'>" & Email_Cargo & "</font></td>" &_
							"</tr>" &_
							"<tr>" &_
							"<td><font size='2' face='verdana'><strong>" & textos_array(7)(2) & "</strong></font></td>" &_
							"<td><font size='2' face='verdana'>" & Email_Depto & "</font></td>" &_
							"</tr>"
		End If
	End If
	If Idioma = "1" Then
		If ID_Formulario <> 4 Then 
			email_confirmacao = email_confirmacao &_
							"<tr>" &_
							"<td><font size='2' face='verdana'><strong>" & textos_array(8)(2) & "</strong></font></td>" &_
							"<td><font size='2' face='verdana'>" & CNPJMask & "</font></td>" &_
							"</tr>"
		End If
	End If
	If ID_Formulario <> 4 Then
		email_confirmacao = email_confirmacao &_
						"<tr>" &_
						"<td><font size='2' face='verdana'><strong>" & textos_array(9)(2) & "</strong></font></td>" &_
						"<td><font size='2' face='verdana'>" & Razao & "</font></td>" &_
						"</tr>"
	End If
	
	'Exibir LINK para Compra do TICKET
	Dim AtivarEComerce : AtivarEComerce = False
	If ID_Edicao = "46" and idioma = "1" And AtivarEComerce Then
		email_confirmacao = email_confirmacao &_
			"<tr>" &_
			"<td colspan='2' align='center'>" &_
			"	<a href='http://credenciamento.btsinforma.com.br/tickets.asp' target='_blank' style='text-decoration:none;color:#00F;'>" &_
			"		<font size='2' face='verdana'>" &_
			"			<strong>Compre aqui seu ingresso!</strong>" &_
			"		</font>" &_
			"	</a>" &_
			"</td>" &_
			"</tr>"
	End If
	
	email_confirmacao = email_confirmacao &_
						"</table>" &_
						"</td>" &_
						"</tr>" &_
						"<tr>" &_
						"<td>&nbsp;</td>" &_
						"</tr>" &_
						"<tr>" &_
						"<td><div align='center'><img src='http://credenciamento.btsinforma.com.br" & RS_Edicoes_Configuracao("Logo_Email") & "'/></div></td>" &_
						"</tr>" &_
						"<tr>" &_
						"<td>&nbsp;</td>" &_
						"</tr>" &_
						"<tr>" &_
						"<td><font size='2' face='verdana'>" &_
						"" & texto_rodape & "</td>" &_
						"</font></tr>" &_
						"<tr>" &_
						"<td>&nbsp;</td>" &_
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


			html = email_confirmacao
'			assunto = "Cad. " & Nome_Formulario & " - " & RS_Edicoes_Configuracao("Feira") & " " & RS_Edicoes_Configuracao("Ano") & " - Credenciamento OnLine BTS Informa"
			assunto = RS_Edicoes_Configuracao("Feira") & " " & RS_Edicoes_Configuracao("Ano") & " - Credenciamento OnLine BTS Informa"

			'response.write(html)
			'response.end()

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
        <br><div style="background-color:#000; color:#FFF;">{ obs: <%=observacao%> - Email enviado para: <%=email%> }</div>
        <%
	End If 
	
	On Error GoTo 0
	
End Function 
%>