<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%> 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Credenciamento <%=Year(Now())%> - BTS Informa</title>
<link href="/css/estilos.css" rel="stylesheet" type="text/css">
</head>
<body style="color:#666">
<%
response.Buffer = True
response.Expires = -1
response.AddHeader "Cache-Control", "no-cache"
response.AddHeader "Pragma", "no-cache" 
%>
<!--#include virtual="/admin/inc/gravar_limpar_texto.asp"-->
<!--#include virtual="/admin/inc/acentos_htm.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<!--#include virtual="/scripts/enviar_email.asp"-->
<%
response.Charset = "utf-8" 
response.ContentType = "text/html" 

'=======================================================================
'	Select 
'		ID_Formulario
'		,Nome
'	FROM Formularios
' **** Resultado
'	ID_Formulario - Nome
'	1 - Empresa
'	2 - Entidades
'	3 - Imprensa
'	4 - Pessoa Física
'	5 - Universidades
'	6 - Alunos

ID_Formulario	=	5 ' Universidades
'=======================================================================

'	For Each item In Request.Form
'		Response.Write "" & item & " - Value: " & Request.Form(item) & "<BR />"
'	Next

Function fixInclude(content)
   out = ""

   'position for  aspStartTag
   pos1 = instr(content,"<%")

   'position for aspEndTag
   pos2 = instr(content,"%"& ">")

   'if there exists aspStartTag
   if pos1 > 0 then

      'text content  before aspStartTag
      before = mid(content,1,pos1-1)

      'remove linebreaks
      before = replace(before,vbcrlf,"")

      'put content into a response.write
      before = vbcrlf & "response.write "" " & before & """" &vbcrlf

      'get code content between aspStartTag  and  aspEndTag
      middle = mid(content,pos1+2,(pos2-pos1-2))

      'get text content after aspEndTag
      after = mid(content,pos2+2,len(content))

      'recurse through the rest
      out = before & middle & fixInclude(after)

   'did not find any aspStartTags
   else
      'remove linebreaks in file
      content = replace(content,vbcrlf,"")
      out = vbcrlf & "response.write""" & content & """"
   end if

   fixInclude = out
End Function

Function getMappedFileAsString(byVal strFilename)
	Const ForReading = 1
	
	Dim fso
	Set fso = Server.CreateObject("Scripting.FilesystemObject")
	
	Dim ts
	Set ts = fso.OpenTextFile(Server.MapPath(strFilename), ForReading)
	
	script = ts.ReadAll
	script = fixInclude(script)
	getMappedFileAsString = script
	ts.close
	
	Set ts = nothing
	Set fso = Nothing
End Function

	' Ler o Arquivo e Executar
	Execute getMappedFileAsString("/scripts/tratar_campos.asp")

'=======================================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'=======================================================================

'****************************
' Descricao dos PASSOS
' 1) Utilizar Campos HIDDEN
' 2) Verificar se a Empresa EXISTE
'	A) Caso SIM - Atualizar
'	B) Senão 	- Cadastrar
' 3) Verificar se o Visitante EXISTE
'	A) Caso SIM	- Atualizar
'	B) Senão	- Cadastrar
' 4) Gravar Relacionamento CADASTRO
' 5) Disparar email de confirmação
' 5) Postar Info´s para 
'****************************

	%>
    <div style="background-color:#FF0;">
    	&bull; ID_Formulario: <%=ID_Formulario%><br>
    	&bull; Origem CNPJ: <%=origem_cnpj%><br>
        &bull; Empresa: <%=id_empresa%><br>
        &bull; Origem CPF: <%=origem_cpf%><br>
        &bull; Visitante: <%=id_visitante%><br>
    </div>
	<%

'=======================================================================
	'2) B - [CADASTRAR] - Empresa veio do banco anterior ou não existe
	If (Trim(origem_cnpj) = "" OR Lcase(origem_cnpj) = "empresa_old") AND Len(Trim(id_empresa)) = 0 Then
		response.write("&bull; - CADASTRAR EMPRESA<br>")
		Execute getMappedFileAsString("/scripts/cadastrar_empresa.asp")
		id_empresa = Novo_ID_Empresa
	End If
'=======================================================================
	'2) A - [ATUALIZAR] - Empresa existe no banco Atual
	If UCase(origem_cnpj) = "NEW" AND Len(id_empresa) > 0 Then
		response.write("&bull; - ATUALIZAR EMPRESA<br>")
		Execute getMappedFileAsString("/scripts/atualizar_empresa.asp")
	End If
'=======================================================================
	'3) B - [CADASTRAR] - Empresa veio do banco anterior ou não existe
	If (Trim(origem_cpf) = "" OR Lcase(origem_cpf) = "old") AND Len(Trim(id_visitante)) = 0 Then
		response.write("&bull; - CADASTRAR VISITANTE<br>")
		Execute getMappedFileAsString("/scripts/cadastrar_visitante.asp")
		id_visitante = Novo_ID_Visitante
	End If
'=======================================================================
	'3) A - [ATUALIZAR] - Empresa existe no banco Atual
	If Ucase(origem_cpf) = "NEW" AND Len(id_visitante) > 0 Then
		response.write("&bull; - ATUALIZAR VISITANTE<br>")
		Execute getMappedFileAsString("/scripts/atualizar_visitante.asp")
	End If
'=======================================================================

'=======================================================================
' Inserir Produtos Selecionado
Lista_Produtos = Split(produtos_inserir,";")
	For i = Lbound(Lista_Produtos) to Ubound(Lista_Produtos) -1
	response.write("i: " & i & " - v: " & Lista_Produtos(i) & "<br>")
		SQL_Cad_Produto = 	"INSERT INTO Relacionamento_Produto_Edicao_Empresa_Visitante_v2 " &_
						"	(ID_Edicao " &_
						"	,ID_Empresa " &_
						"	,ID_Visitante " &_
						"	,Principal_Produto " &_
						"	,Ativo) " &_
						"VALUES " &_
						"	(" &_
						"	" & id_edicao & " " &_
						"	," & id_empresa & " " &_
						"	," & id_visitante & " " &_				
						"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Trim(Lista_Produtos(i)),150) & "')) " &_
						"	,1 " &_
						"); "

		response.write("<b>SQL_Cad_Produto</b><br>" & SQL_Cad_Produto & "<hr>")
		' Executando Gravação
		Set RS_Cad_Produto = Conexao.Execute(SQL_Cad_Produto)
	Next
'=======================================================================

	' Enviar EMAIL
	Enviar_Email id_edicao, id_idioma, ID_Formulario, Email, Novo_ID_Rel_Cadastro, CPF, Passaporte, Nome, Cargo, Depto, CNPJ, Razao

Conexao.Close

Response.Clear()

Session("cliente_empresa") 	= id_empresa
Session("cliente_logado") 	= True				' Para a tela de Alunos não redirecionar ao LOGIN
%>
<form id="confirmacao" name="confirmacao" method="POST" action="/alunos/cadastrar.asp">
	<input type="hidden" name="id_edicao" 			value="<%=id_edicao%>">
	<input type="hidden" name="id_idioma" 			value="<%=id_idioma%>">
	<input type="hidden" name="id_tipo" 			value="<%=id_tipo%>">
	<input type="hidden" name="frmID_Cadastro" 		value="<%=Novo_ID_Rel_Cadastro%>">
	<input type="hidden" name="frmID_Empresa" 		value="<%=id_empresa%>">
	<input type="hidden" name="frmNome" 			value="<%=Nome%>">
	<input type="hidden" name="frmCPF" 				value="<%=CPF%>">
	<input type="hidden" name="frmPassaporte" 		value="<%=Passaporte%>">
	<input type="hidden" name="frmCargo" 			value="<%=Cargo%>">
	<input type="hidden" name="frmDepartamento" 	value="<%=Depto%>">
	<input type="hidden" name="frmCNPJ" 			value="<%=CNPJ%>">
	<input type="hidden" name="frmRazaoSocial" 		value="<%=Razao%>">
</form>
<script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
<script language="Javascript">
	$(document).ready(function(){
		$("#confirmacao").submit();
	});
</script>
</body>
</html>