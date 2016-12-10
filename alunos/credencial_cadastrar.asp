<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%> 
<!--#include virtual="/admin/inc/gravar_limpar_texto.asp"-->
<%
	'- Verifica se existe no TRADECENTER 
	id_tipo 		= Limpar_Texto(Request("id_tipo"))
	nome 			= Left(Limpar_Texto(Request("nome")),100)
	email 			= Left(Limpar_Texto(Request("email")),100)
	curso			= Left(Limpar_Texto(Request("curso")),40)

'	response.write("{ msg: ' " & email & " ' }")
'	response.end()

	
	If id_tipo = "-" Then 
		response.write("{ msg: 'tipo invalido' }")
	'=======================================================================
	ElseIf Len(nome) < 3 Then
		response.write("{ msg: 'nome curto' }")
	'=======================================================================
	ElseIf Len(email) < 3 Then
		response.write("{ msg: 'e-mail curto email email' }")
	'=======================================================================
	ElseIf Len(curso) < 2 Then
		response.write("{ msg: 'curso curto' }")
	'=======================================================================
	ElseIf Session("cliente_edicao") = "" Then
		response.write("{ msg: 'sessao expirou' }")
		response.Redirect("/?erro=1")
	'=======================================================================
	Else
		Set Conexao_TC = Server.CreateObject("ADODB.Connection")
		Conexao_TC.Open Application("cnn")
		
		SQL 	=	"Select " &_
					"	ID_Universidade_Credencial " &_
					"From Universidade_Credenciais " &_
					"Where " &_
					"	Nome = Upper(dbo.sp_rm_accent_pt_latin1('" & nome & "')) " &_
					"	AND Curso = Upper(dbo.sp_rm_accent_pt_latin1('" & Curso & "')) " &_
					"	AND ID_Universidade_TipoCredencial = " & id_tipo & " " &_
					"	AND ID_Edicao = " & Session("cliente_edicao") & " " &_
					"	AND ID_Empresa = " & Session("cliente_empresa") & " "
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		RS.Open SQL, Conexao_TC
		'=======================================================================
		If RS.BOF or RS.EOF Then
		'=======================================================================
			SQL_Qtde_Preenchida =   "Select " &_
									"   Count(ID_Universidade_Credencial) as Total " &_
									"From Universidade_Credenciais " &_
									"Where  " &_
									"   ID_Edicao = " & Session("cliente_edicao") & " " &_
									"   AND ID_Empresa = " & Session("cliente_empresa") & " "

			Set RS_Qtde_Preenchida = Server.CreateObject("ADODB.Recordset")
			RS_Qtde_Preenchida.Open SQL_Qtde_Preenchida, Conexao_TC
			
			Qtde_Preenchida = 0
			If not RS_Qtde_Preenchida.BOF or not RS_Qtde_Preenchida.EOF Then
				Qtde_Preenchida = RS_Qtde_Preenchida("total")
				If isNull(Qtde_Preenchida) Then Qtde_Preenchida = 0
				RS_Qtde_Preenchida.Close
			End If		
			'=======================================================================
			
			qtde_disponivel = 99
			If Cint( qtde_disponivel ) > Qtde_Preenchida Then
				SQL_Inserir = 	"Insert Into Universidade_Credenciais " &_
								"(ID_Universidade_TipoCredencial, nome, Email, Curso, id_edicao, id_empresa) " &_
								"values " &_
								"(" & id_tipo & ", Upper(dbo.sp_rm_accent_pt_latin1('" & nome & "')), Upper(dbo.sp_rm_accent_pt_latin1('" & email & "')), Upper(dbo.sp_rm_accent_pt_latin1('" & curso & "')), " & Session("cliente_edicao") & ", " & Session("cliente_empresa") & ")"

				Set RS_Inserir = Server.CreateObject("ADODB.Recordset")
				RS_Inserir.Open SQL_Inserir, Conexao_TC
				response.write("{ msg: 'cadastrada', qtde_restante: '" & Cint( qtde_disponivel ) - Qtde_Preenchida - 1 & "', qtde_preenchida : '" & Qtde_Preenchida + 1 & "' }")
				'Session("cliente_msg") = "Credencial cadastrada !<br>Restam <b>" & Cint( (Session("cliente_m2") * 0.4) + Itens_Solicitados) - 1 & "</b> para preencher."
			Else
				response.write("{ msg: 'qtde esgotou' }")
			End If
		'=======================================================================
		Else		
			response.write("{ msg: 'duplicada' }")
			RS.Close
		End If
		'=======================================================================
		Conexao_TC.Close
	End If
%>