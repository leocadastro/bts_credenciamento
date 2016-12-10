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
ID_Cargo = Trim(Limpar_Texto(Request("ID_Cargo")))
Idioma 	 = Session("cliente_idioma")

If ID_Cargo = "-" Then ID_Cargo = 0

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
If Len(ID_Cargo) <> 0 Then

	'=======================================================================
	' Verificando cadastro no banco novo Credenciamento_2012
	'=======================================================================
	SQL_SubCargo = 		"SELECT " &_
						"	S.ID_SubCargo, " &_
						"	S.SubCargo_" & SgIdioma & " as Subcargo " &_
						"FROM Relacionamento_Edicoes_SubCargo as RES " &_
						"Inner Join SubCargo as S ON S.ID_SubCargo = RES.ID_SubCargo " &_
						"WHERE " &_
						"	S.Ativo = 1 " &_
						"	AND RES.ID_Edicao = " & Session("cliente_edicao") & " " &_
						"	AND ID_Cargo = " & ID_Cargo & " " &_
						"ORDER BY Subcargo "

	'response.Write("<strong>SQL_SubCargo</strong><hr>" & SQL_SubCargo & "<hr>")
	Set RS_Verificar = Server.CreateObject("ADODB.Recordset")
	RS_Verificar.Open SQL_SubCargo, Conexao

	'=======================================================================
	' Verificando o Retorno da Query SQL_Verificar
	'=======================================================================
	If RS_Verificar.BOF or RS_Verificar.EOF Then

				%>
                { 
                Resultado : '0',
                ResultadoTXT : 'Cadastro não localizado - Cargo'
                }
                <%

	Else

		'=======================================================================
		' Cadastro localizado no Banco Novo Credenciamento_2012
		' TRATANDO AS VARIAVEIS
		'=======================================================================

		%>
		{ "SubCargos":
			[
		<%


		While not RS_Verificar.EOF
			If RS_Verificar("Subcargo") = "Outros" or RS_Verificar("Subcargo") = "Others" or RS_Verificar("Subcargo") = "Other" or RS_Verificar("Subcargo") = "Otros" Then
				OptionSubCargo = " { ""id"": """ & RS_Verificar("ID_SubCargo") & """, ""nome"": """ & Trim(caixaAltaBaixa("caixa_altabaixa",RS_Verificar("Subcargo"))) & """ } "
				OptionSubCargoOutros = true
			Else
				SubCargo = Subcargo & " { ""id"": """ & RS_Verificar("ID_SubCargo") & """, ""nome"": """ & Trim(caixaAltaBaixa("caixa_altabaixa",RS_Verificar("Subcargo"))) & """ }"
				OptionSubCargoOutros = false
			End If
			RS_Verificar.MoveNext()
			
			If not RS_Verificar.EOF Then
				If OptionSubCargoOutros = false then
					SubCargo = SubCargo & ","
				end if
			End If
		Wend
		RS_Verificar.Close	
		
		If SubCargo <> "" Then
			SubCargo = SubCargo & OptionSubCargo
			response.write(SubCargo)
		Else
			response.write(OptionSubCargo)
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