<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%> 
<!--#include virtual="/includes/limpar_texto.asp"-->
<%
	'- Verifica se existe no TRADECENTER 
	id 		 		= Limpar_Texto(Request("id"))
	id_tipo 		= Limpar_Texto(Request("id_tipo"))
	nome 			= Left(Limpar_Texto(Request("nome")),100)
	curso			= Left(Limpar_Texto(Request("curso")),40)

	If id_tipo = "-" Then 
		response.write("{ msg: 'tipo invalido' }")
	'=======================================================================
	ElseIf Len(nome) < 3 Then
		response.write("{ msg: 'nome curto' }")
	'=======================================================================
	ElseIf Len(curso) < 3 Then
		response.write("{ msg: 'curso curto' }")
	'=======================================================================
	ElseIf Session("cliente_edicao") = "" Then
		response.write("{ msg: 'sessao expirou' }")
		response.Redirect("/?erro=1")
	'=======================================================================
	Else
		Set Conexao_TC = Server.CreateObject("ADODB.Connection")
		Conexao_TC.Open Application("cnn")
		
		SQL_Atualizar =	"Update Universidade_Credenciais " &_
						"Set " &_
						"	ID_Universidade_TipoCredencial = " & id_tipo & ", " &_ 
						"	nome = Upper(dbo.sp_rm_accent_pt_latin1('" & nome & "')), " &_
						"	curso = Upper(dbo.sp_rm_accent_pt_latin1('" & curso & "')) " &_
						"Where ID_Universidade_Credencial = " & id

'response.write(SQL_Atualizar)

		Set RS_Atualizar = Server.CreateObject("ADODB.Recordset")
		RS_Atualizar.Open SQL_Atualizar, Conexao_TC
		response.write("{ msg: 'atualizada'}")
		'=======================================================================
		Conexao_TC.Close
	End If
%>