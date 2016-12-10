<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/includes/limpar_texto.asp"-->
<%

idioma 	= Limpar_Texto(Request("i"))
edicao 	= Limpar_Texto(Request("e"))

If Len(Trim(idioma)) > 0 AND Len(Trim(edicao)) > 0 Then
'===========================================================
	Set Conexao = Server.CreateObject("ADODB.Connection")
	Conexao.Open Application("cnn")
'===========================================================
	' Listagem de Tipos deste IDIOMA e EDICAO
	SQL_Tipos	= 	"Select " &_
					"	Tc.ID_Tipo_Credenciamento, " &_
					"	Tc.Nome, " &_
					"	Tc.IMG_Box, " &_
					"	Case " &_
					"		When Et.URL_Especial is not Null Then Et.URL_Especial " &_
					"		Else Tc.URL " &_
					"	End " &_
					"	as URL, "&_
					"	Tc.ID_Formulario " &_
					"From Edicoes_Tipo as Et   " &_
					"Inner Join Eventos_Edicoes as Ee ON Ee.ID_Edicao = Et.ID_Edicao   " &_
					"Inner Join Edicoes_Configuracao as Ec ON Ec.ID_Edicao = Et.ID_Edicao  " &_
					"Inner Join Eventos as E ON E.ID_Evento = Ee.ID_Evento   " &_
					"Inner Join Tipo_Credenciamento as Tc ON Tc.ID_Tipo_Credenciamento = Et.ID_Tipo_Credenciamento   " &_
					"Where  " &_
					"	Ee.Ativo = 1  " &_
					"	AND Ec.Ativo = 1  " &_
					"	AND E.Ativo = 1  " &_
					"	AND Tc.Ativo = 1  " &_
					"	AND	Tc.ID_Idioma = " & idioma & " " &_
					"	AND	Et.ID_Edicao = " & edicao & " " &_
					"	AND Et.Ativo = 1 " &_
					"	AND getDate() >= Inicio " &_
					"	AND getDate() <= Fim " &_
					"Order by Tc.Nome "

					' Alteração pro ambiente de TESTE
'					
' ! ATENÇÃO ==============================================================

	
	Set RS_Tipos = Server.CreateObject("ADODB.Recordset")
	RS_Tipos.CursorType = 0
	RS_Tipos.LockType = 1
	RS_Tipos.Open SQL_Tipos, Conexao, 1

'===========================================================
	If not RS_Tipos.BOF or not RS_Tipos.EOF Then
		lista = "["
		While not RS_Tipos.EOF
			nome		= RS_Tipos("nome")
			img			= RS_Tipos("img_box")
			url			= Trim(RS_Tipos("url"))
			tipo		= RS_Tipos("ID_Tipo_Credenciamento")
			formulario 	= RS_Tipos("ID_Formulario")
			lista 	= lista & " { nome: '" & nome & "', img: '" & img & "', url: '" & url & "', tipo: '" & tipo & "', formulario: '" & formulario & "' }" 
			RS_Tipos.MoveNext
			If not RS_Tipos.EOF Then 
				lista = lista & ", "
			Else
				lista = lista & "]"
			End If
		Wend
		RS_Tipos.Close
		response.write("{ msg: 'ok', itens: " & lista & " }")
	Else
		response.write("{ msg: 'sem tipos disponiveis' }")
	End If
'===========================================================
	Conexao.Close
'===========================================================
Else
	response.write("{ msg: 'ids nao recebidos' }")
End IF
%>