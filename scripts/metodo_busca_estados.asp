<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%> 
<%
Response.Expires = -1
Response.AddHeader "Cache-Control", "no-cache"
Response.AddHeader "Pragma", "no-cache" 
%>
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/acentos_htm.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
response.Charset = "utf-8" 
response.ContentType = "text/html" 

If Session("cliente_edicao") = "" OR Session("cliente_idioma") = "" OR Session("cliente_tipo") = "" Then
	response.Redirect("/?erro=1")
End If

'=======================================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
			  Conexao.Open Application("cnn")
'=======================================================================

'=======================================================================
' Pegando o CPF para comparar com o Banco novo e com o Banco Antigo
'=======================================================================
Idioma 	 = Session("cliente_idioma")

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

	'=======================================================================
	' Verificando cadastro no banco novo Credenciamento_2012
	'=======================================================================
	' Select de Estados
	SQL_Estado = 		"SELECT " &_
						"	ID_UF, " &_ 
						"	Sigla, " &_ 
						"	Estado " &_
						"FROM UF " &_
						"WHERE " &_
						"	Ativo = 1 " &_
						"ORDER BY Estado "
	'response.write(SQL_Estado)
	Set RS_Verificar = Server.CreateObject("ADODB.Recordset")
	RS_Verificar.CursorType = 0
	RS_Verificar.LockType = 1
	RS_Verificar.Open SQL_Estado, Conexao

	'=======================================================================
	' Verificando o Retorno da Query SQL_Verificar
	'=======================================================================
	If RS_Verificar.BOF or RS_Verificar.EOF Then

				%>
                { 
                Resultado : '0',
                ResultadoTXT : 'Estado não localizado'
                }
                <%

	Else

		'=======================================================================
		' Cadastro localizado no Banco Novo Credenciamento_2012
		' TRATANDO AS VARIAVEIS
		'=======================================================================

		%>
		{ "Estados":
			[
		<%

		While not RS_Verificar.EOF

				Estado = Estado & " { ""id"": """ & RS_Verificar("ID_UF") & """, ""sigla"": """ & RS_Verificar("Sigla") & """, ""nome"": """ & Trim(caixaAltaBaixa("caixa_altabaixa",RS_Verificar("Estado"))) & """ } "
			
			RS_Verificar.MoveNext()
			
			If not RS_Verificar.EOF Then
				Estado = Estado & ","
			End If
		Wend

		RS_Verificar.Close	
		
		response.write(Estado)

		%>
			],
		 "Resultado" : "1"		
		}
		<%
	End If	
Conexao.Close
%>