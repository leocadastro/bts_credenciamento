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
CNPJOld = Trim(Limpar_Texto(Request("cnpj")))
CNPJ 	= Trim(Limpar_Texto(Request("cnpj")))
CNPJ 	= Replace(CNPJ,".","")
CNPJ 	= Replace(CNPJ,"-","")
CNPJ 	= Replace(CNPJ,"/","")

'=======================================================================
' Verificando valor do campo documento
'=======================================================================
If Len(CNPJ) <> 0 Then

	'=======================================================================
	' Verificando cadastro no banco novo Credenciamento_2012
	'=======================================================================
	SQL_Recuperar_Senha =	"Select Top 1 " &_
							"	E.CNPJ " &_
							"	,E.Senha " &_
							"	,E.Razao_Social " &_
							"	,V.Email " &_
							"	,V.Nome_Credencial " &_
							"From Relacionamento_Cadastro as RC " &_
							"Inner Join Tipo_Credenciamento as TC ON TC.ID_Tipo_Credenciamento = RC.ID_Tipo_Credenciamento " &_
							"Inner Join Empresas as E ON E.ID_Empresa = RC.ID_Empresa " &_
							"Inner Join Visitantes as V ON V.ID_Visitante = RC.ID_Visitante " &_
							"Where  " &_
							"	RC.ID_Tipo_Credenciamento = 13 							/* Alunos */ " &_
							"	AND RC.ID_Edicao = " & Session("cliente_edicao") & "	/* Edição	*/ " &_
							"	AND TC.ID_Idioma = " & Session("cliente_idioma") & "	/* Idioma	*/ " &_
							"	AND E.CNPJ = '" & CNPJ & "'								/* CNPJ	*/ 	   "
	
	'response.Write("<strong>SQL_Recuperar_Senha</strong><hr>" & SQL_Recuperar_Senha & "<hr>")
	
	Set RS_Recuperar_Senha = Server.CreateObject("ADODB.Recordset")
	RS_Recuperar_Senha.CursorType = 0
	RS_Recuperar_Senha.LockType = 1
	RS_Recuperar_Senha.Open SQL_Recuperar_Senha, Conexao
	
	If RS_Recuperar_Senha.BOF or RS_Recuperar_Senha.EOF Then
		%>{ 'retorno' : 'cnpj nao cadastrado' }<%	
	Else
		Razao_Social 	= RS_Recuperar_Senha("Razao_Social")
		Email 			= RS_Recuperar_Senha("Email")
		Nome 			= RS_Recuperar_Senha("Nome_Credencial")
		Senha 			= RS_Recuperar_Senha("Senha")
		
		Tipo			= "Recuperar_Senha"
		
		If Request("admin") = "sim" Then
			Response.Write(Tipo)
			'Response.End()
		End If
		
		Enviar_Email_Senha Session("cliente_edicao"), Session("cliente_idioma"), CNPJ, Razao_Social, Email, Nome, Senha, Tipo, Pedido, Var_Visitante
			
		%>{ 'retorno' : 'email enviado', 'razao' : '<%=razao_social%>', 'email' : '<%=email%>', 'nome' : '<%=nome%>' }<%	
	End If
Else 
	%>{ 'retorno' : 'cnpj invalido' }<%
End If

Conexao.Close
%>