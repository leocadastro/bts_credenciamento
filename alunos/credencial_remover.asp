<!--#include virtual="/includes/limpar_texto.asp"-->
<%
	'- Verifica se existe no TRADECENTER 
	id 		= Limpar_Texto(Request("id"))
	
	Set Conexao_TC = Server.CreateObject("ADODB.Connection")
	Conexao_TC.Open Application("cnn")
	
	SQL 	=	"Delete From Universidade_Credenciais " &_
				"Where " &_
				"	ID_Universidade_Credencial = " & id
	
	Set RS = Server.CreateObject("ADODB.Recordset")
	RS.Open SQL, Conexao_TC
	
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
		RS_Qtde_Preenchida.Close
	End If

	response.write("{ msg: 'ok', qtde_restante: '" & Cint(99) - Qtde_Preenchida & "', qtde_preenchida : '" & Qtde_Preenchida & "' }")

	Conexao_TC.Close
%>