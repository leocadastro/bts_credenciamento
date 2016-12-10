<!--#include virtual="/includes/limpar_texto.asp"-->
<!--#include virtual="/includes/texto_caixaAltaBaixa.asp"-->
<%
'===========================================================
Qs = Request.ServerVariables("QUERY_STRING")
'===========================================================
idioma = Session("idioma")
'===========================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'===========================================================
	SQL_Qtde_Preenchida =	"Select " &_
							"	Count(ID_Universidade_Credencial) as Total " &_
							"From Universidade_Credenciais " &_
							"Where  " &_
							"	ID_Edicao = " & Session("cliente_edicao") & " " &_
							"	AND ID_Empresa = " & Session("cliente_empresa") & " "
	Set RS_Qtde_Preenchida = Server.CreateObject("ADODB.Recordset")
	RS_Qtde_Preenchida.Open SQL_Qtde_Preenchida, Conexao
	
	Qtde_Preenchida = 0
	If not RS_Qtde_Preenchida.BOF or not RS_Qtde_Preenchida.EOF Then
		Qtde_Preenchida = RS_Qtde_Preenchida("total")
		RS_Qtde_Preenchida.Close
	End If
'===========================================================

' Feira Selecionada
feira_logo = Session("cliente_Logo_Faixa")
feira_fundo = Session("cliente_Bg_Faixa")
'===========================================================
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Credenciamento - BTS Informa</title>
<link href="/css/estilos.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="770" border="0" cellpadding="0" cellspacing="2">
  <tr>
	<td height="20" colspan="5" align="left" class="arial fs_13px cor_branco " style="padding-left:5px;" background="<%=feira_fundo%>"><strong>Credenciais Cadastradas</strong></td>
  </tr>
  <tr>
	<td width="50" height="20" align="center" class="arial fs_13px cor_cinza2 borda_tabela" ><b>Nº </b></td>
	<td width="75" height="20" align="center" class="arial fs_13px cor_cinza2 borda_tabela"><b>Tipo</b></td>
	<td height="20" align="center" class="arial fs_13px cor_cinza2 borda_tabela"><b>Nome</b></td>
	<td height="20" align="center" class="arial fs_13px cor_cinza2 borda_tabela"><b>E-mail</b></td>
	<td height="20" align="center" class="arial fs_13px cor_cinza2 borda_tabela"><b>Curso</b></td>
	<td width="50" height="20" align="center" class="arial fs_13px cor_cinza2 borda_tabela" ><b>Editar</b></td>
	<td width="50" height="20" align="center" class="arial fs_13px cor_cinza2 borda_tabela" ><b>Remover</b></td>
  </tr>
  <%
	SQL_Credenciais_Preenchidas = 	"Select " &_
									"	ID_Universidade_Credencial " &_
									"	,T.ID_Universidade_TipoCredencial " &_
									"	,T.Tipo " &_
									"	,Nome " &_
									"	,Email" &_
									"	,Curso " &_
									"From Universidade_Credenciais as U " &_
									"Inner Join Universidade_TipoCredencial as T ON T.ID_Universidade_TipoCredencial = U.ID_Universidade_TipoCredencial " &_
									"Where " &_
									"	ID_Edicao = " & Session("cliente_edicao") & " " &_
									"	AND ID_Empresa = " & Session("cliente_empresa") & " " &_
									"Order by ID_Universidade_Credencial DESC "
									
	Set RS_Credenciais_Preenchidas = Server.CreateObject("ADODB.Recordset")
	RS_Credenciais_Preenchidas.Open SQL_Credenciais_Preenchidas, Conexao
	
	n = Qtde_Preenchida
	If not RS_Credenciais_Preenchidas.BOF or not RS_Credenciais_Preenchidas.EOF Then
		While not RS_Credenciais_Preenchidas.EOF
			ID		= RS_Credenciais_Preenchidas("ID_Universidade_Credencial")
			ID_Tipo = RS_Credenciais_Preenchidas("ID_Universidade_TipoCredencial")
			Tipo	= RS_Credenciais_Preenchidas("Tipo")
			Nome	= RS_Credenciais_Preenchidas("Nome")
			Email	= RS_Credenciais_Preenchidas("Email")
			Curso	= RS_Credenciais_Preenchidas("Curso")
  %>
  <tr>
	<td align="center" class="arial fs_12px cor_cinza1 borda_tabela"><%=preencher_zeros(4,n)%></td>
	<td height="30" align="left" style="padding-left:10px;" bgcolor="#e4e5e6" class="arial fs_12px cor_cinza2 b borda_tabela"><%=Tipo%></td>
	<td height="30" align="left" style="padding-left:10px;" bgcolor="#e4e5e6" class="arial fs_12px cor_cinza2 b borda_tabela"><%=Nome%></td>
	<td height="30" align="left" style="padding-left:10px;" bgcolor="#e4e5e6" class="arial fs_12px cor_cinza2 b borda_tabela"><%=Email%></td>
	<td height="30" align="left" style="padding-left:10px;" bgcolor="#e4e5e6" class="arial fs_12px cor_cinza2 b borda_tabela"><%=Curso%></td>
	<td align="center" class="arial fs_12px cor_cinza1 borda_tabela">
        <img class="cursor" onClick="top.atualizar_credencial(<%=ID%>,'<%=nome%>','<%=Email%>','<%=Curso%>','<%=ID_Tipo%>');"
        src="/img/geral/icones/ico_editar.gif" width="20" height="20">
	</td>
	<td align="center" class="arial fs_12px cor_cinza1 borda_tabela">
		<img class="cursor" onClick="top.remover_credencial(<%=ID%>,'<%=nome%>','<%=email%>','<%=Curso%>');"
		src="/img/geral/icones/nok.gif" width="20" height="20">
	</td>
  </tr>
  <%
			n = n - 1
			RS_Credenciais_Preenchidas.MoveNext
		Wend
	Else
  %>
  <tr>
	<td colspan="7" height="30" align="center" bgcolor="#990000" class="arial fs_12px cor_branco b borda_tabela">Nenhuma Credencial Cadastrada</td>
  </tr>
  <%
	End If
  %>
  </table>
</body>
</html>