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
ID_Ramo 	= Trim(Limpar_Texto(Request("ID_Ramo")))
Idioma 	 	= Session("cliente_idioma")

If ID_Ramo = "-" Then ID_Cargo = 0

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
' Verificando valor do campo documento
'=======================================================================
If Len(ID_Ramo) <> 0 Then

	'=======================================================================
	' Verificando cadastro no banco novo Credenciamento_2012
	'=======================================================================
	SQL_Atividade = 	"SELECT " &_
						"	A.ID_Atividade, " &_
						"	A.Atividade_" & SgIdioma & " as Atividade " &_
						"FROM Relacionamento_Edicoes_Atividade as REA " &_
						"Inner Join AtividadeEconomica as A ON A.ID_Atividade = REA.ID_Atividade " &_
						"WHERE " &_
						"	A.Ativo = 1 " &_
						"	AND REA.ID_Edicao = " & Session("cliente_edicao") & " " &_
						"	AND A.ID_Ramo = " & ID_Ramo & " " &_
						"ORDER BY Atividade " 

	'response.Write("<strong>SQL_Atividade</strong><hr>" & SQL_Atividade & "<hr>")
	Set RS_Verificar = Server.CreateObject("ADODB.Recordset")
	RS_Verificar.Open SQL_Atividade, Conexao, 1

	'=======================================================================
	' Verificando o Retorno da Query SQL_Verificar
	'=======================================================================
	If RS_Verificar.BOF or RS_Verificar.EOF Then

				%>
                { 
                Resultado : '0',
                ResultadoTXT : 'Cadastro não localizado - Atividade'
                }
                <%

	Else

		'=======================================================================
		' Cadastro localizado no Banco Novo Credenciamento_2012
		' TRATANDO AS VARIAVEIS
		'=======================================================================

		%>
		{ "Atividades":
			[
		<%

		While not RS_Verificar.EOF
			
			If RS_Verificar("Atividade") = "Outros" or RS_Verificar("Atividade") = "Others" or RS_Verificar("Atividade") = "Otros" Then
				OptionAtividades = " { ""id"": """ & RS_Verificar("ID_Atividade") & """, ""nome"": """ & Trim(caixaAltaBaixa("caixa_altabaixa",RS_Verificar("Atividade"))) & """ } "
			Else
				Atividades = Atividades & " { ""id"": """ & RS_Verificar("ID_Atividade") & """, ""nome"": """ & Trim(caixaAltaBaixa("caixa_altabaixa",RS_Verificar("Atividade"))) & """ }, "
			End If

			RS_Verificar.MoveNext()

			If RS_Verificar.EOF Then
				If OptionAtividades <> "" then
					Atividades = Atividades
				End if
			End If
		Wend

		RS_Verificar.Close	

		If Atividades <> "" Then
			Atividades = Atividades & OptionAtividades
			response.write(Atividades)
		Else
			response.write(OptionAtividades)
		End If

		%>
			],
		 "Resultado" : "1"		
		}
		<%
	End If	
Else 

End If
Conexao.Close
%>