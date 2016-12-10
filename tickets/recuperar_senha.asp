<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%> 
<%
Response.Expires = -1
Response.AddHeader "Cache-Control", "no-cache"
Response.AddHeader "Pragma", "no-cache" 
%>
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/acentos_htm.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<!--#include virtual="/scripts/enviar_email_senha.asp"-->
<%
response.Charset = "utf-8" 
response.ContentType = "text/html" 

'=======================================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
			  Conexao.Open Application("cnn")
'=======================================================================

'=======================================================================
' Pegando o CNPJ para comparar com o Banco novo e com o Banco Antigo
'=======================================================================
login 	= Trim(Limpar_Texto(Request("login")))
login 	= Replace(login,".","")
login 	= Replace(login,"-","")
login 	= Replace(login,"/","")

If IsNumeric(login) then
	login_id = login
Else
	login_id = 0
End If

'=======================================================================
' Verificando valor do campo documento
'=======================================================================
If Len(login) <> 0 Then

	'=======================================================================
	' Verificando cadastro no banco novo Credenciamento_2012
	'=======================================================================
	SQL_Recuperar_Senha =	"Select " &_
							"	Top 1 " &_
							"	V.ID_Visitante, " &_
							"	V.Nome_Completo, " &_
							"	V.Email, " &_
							"	V.Senha " &_
							"From Visitantes V " &_
							"Inner Join Relacionamento_Cadastro R " &_
							"	On R.ID_Visitante=V.ID_Visitante " &_
							"Where " &_
							"	R.ID_Edicao = '" & session("cliente_edicao") & "'  " &_
							"	AND (V.ID_Visitante = '" & login_id & "'  " &_
							"	or V.CPF = '" & login & "' " &_
							"	or V.Passaporte = '" & login & "' ) " &_
							"Order by V.Data_Atualizacao DESC, V.Data_Cadastro DESC"
	
	
	Set RS_Recuperar_Senha = Server.CreateObject("ADODB.Recordset")
	RS_Recuperar_Senha.Open SQL_Recuperar_Senha, Conexao, 3, 3
	
	If RS_Recuperar_Senha.BOF or RS_Recuperar_Senha.EOF Then
		%>{ 'retorno' : 'login nao cadastrado' }<%	
	Else
		ID_Visitante	= RS_Recuperar_Senha("ID_Visitante")
		Email 			= RS_Recuperar_Senha("Email")
		Nome 			= RS_Recuperar_Senha("Nome_Completo")
		Senha 			= RS_Recuperar_Senha("Senha")
	
		' Se a senha for VAZIA, altere para o ID_REL_CAD
		If Len(Trim(Senha)) = 0 Or IsNull(Senha) Then

			SQL_Get_ID_RELCAD = 	"Select " &_
									"	top 1 " &_
									"	ID_Relacionamento_Cadastro " &_
									"From Relacionamento_Cadastro " &_
									"Where " &_
									"	ID_Visitante = " & ID_Visitante & "  " &_
									"Order by ID_Relacionamento_Cadastro DESC"
									
			Set RS_Get_ID_RELCAD = Server.CreateObject("ADODB.Recordset")
				RS_Get_ID_RELCAD.CursorType = 0
				RS_Get_ID_RELCAD.LockType = 1
				RS_Get_ID_RELCAD.Open SQL_Get_ID_RELCAD, Conexao
				
			If not RS_Get_ID_RELCAD.BOF or not RS_Get_ID_RELCAD.EOF Then
				nova_senha = RS_Get_ID_RELCAD("ID_Relacionamento_Cadastro")
				RS_Get_ID_RELCAD.Close
				
				SQL_Atualizar_senha = 	"Update Visitantes " &_
										"Set	Senha = " & nova_senha & " " &_
										"Where ID_Visitante = " & ID_Visitante
																			
				Conexao.Execute(SQL_Atualizar_senha)
				
				senha = nova_senha
			End If 
			
		End If
		
		'response.write ("senha: " & senha & "<br>")
		'response.write ("nova_senha: " & nova_senha & "<br>")
		Enviar_Email_Senha Session("cliente_edicao"), Session("cliente_idioma"), "", "", Email, Nome, Senha, "Recuperar_Senha", "", ""
		
		%>{ 'retorno' : 'email enviado', 'email' : '<%=Lcase(email)%>', 'nome' : '<%=nome%>' }<%
	End If
Else 
	%>{ 'retorno' : 'login invalido' }<%
End If

Conexao.Close
%>