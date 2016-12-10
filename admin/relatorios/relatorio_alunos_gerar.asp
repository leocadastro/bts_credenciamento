<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
'if session("admin_id_usuario") <> "21" then
'	response.Redirect("/admin/relatorios/default.asp")
'end if
' * Dados Paginação
evento = Request("evento")

'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================

SQL_Alunos = 	"Select Distinct " &_
				"	Ev.Nome_PTB As Evento " &_
				"	,Ut.Tipo " &_
				"	,Ee.ID_Edicao " &_
				"	,Ee.Ano " &_
				"	,Uc.Nome " &_
				"	,Uc.Curso " &_
				"	,Uc.Email  " &_
				"	,Uc.Ativo " &_
				"	,Em.ID_Empresa " &_
				"	,Em.Razao_Social " &_
				"	,Em.Nome_Fantasia " &_
				"	,Em.Reitor " &_
				"	,Vi.Nome_Completo " &_
				"	,Ca.Cargo_PTB " &_
				"From Universidade_Credenciais As Uc " &_
				"Inner Join Universidade_TipoCredencial AS Ut " &_
				"	On Ut.ID_Universidade_TipoCredencial = Uc.ID_Universidade_TipoCredencial " &_
				"Inner Join Eventos_Edicoes As Ee " &_
				"	On Ee.ID_Edicao = Uc.ID_Edicao " &_
				"Inner Join Eventos As Ev " &_
				"	On Ev.ID_Evento = Ee.ID_Evento " &_
				"Inner Join Empresas As Em " &_
				"	On Em.ID_Empresa = Uc.ID_Empresa " &_
				"Inner Join Relacionamento_Cadastro As Rc " &_
				"	On  Rc.ID_Empresa = Em.ID_Empresa " &_
				"	And Rc.ID_Edicao = Ee.ID_Edicao " &_
				"	And Rc.ID_Tipo_Credenciamento = 13 " &_
				"Inner Join Visitantes As Vi " &_
				"	On Vi.ID_Visitante = Rc.ID_Visitante " &_
				"Inner Join Cargo As Ca " &_
				"	On Ca.ID_Cargo = Vi.ID_Cargo " &_
				"Where " &_
				"	Ee.ID_Edicao = " & Cint(evento) & " " &_
				"Group By " &_
				"	Ev.Nome_PTB " &_
				"	,Ut.Tipo " &_
				"	,Ee.ID_Edicao " &_
				"	,Ee.Ano " &_
				"	,Uc.Nome " &_
				"	,Uc.Curso " &_
				"	,Uc.Email  " &_
				"	,Uc.Ativo " &_
				"	,Em.ID_Empresa " &_
				"	,Em.Razao_Social " &_
				"	,Em.Nome_Fantasia " &_
				"	,Em.Reitor " &_
				"	,Vi.Nome_Completo " &_
				"	,Ca.Cargo_PTB " &_
				"Order By Ev.Nome_PTB, Uc.Nome"
Set RS_Alunos = Server.CreateObject("ADODB.Recordset")
RS_Alunos.Open SQL_Alunos, Conexao

If Not RS_Alunos.EOF Then

	d = Day(Now)
	m = Month(Now)
	a = Year(Now)
	h = Hour(Now)
	n = Minute(Now)
	s = Second(Now)
	If Len(d) < 2 Then d = "0" & d
	If Len(m) < 2 Then m = "0" & m
	If Len(h) < 2 Then h = "0" & h
	If Len(n) < 2 Then n = "0" & n
	If Len(s) < 2 Then s = "0" & s
	data = d & "-" & m & "-" & a & "_" & h & "-" & n & "-" & s
	'FIM - MONTA DATA PARA GERAR ARQUIVO
	
	'INICIO - GERAR ARQUIVO
	formato = "xls"
	arquivo = "REL_ALUNOS_"
	extensao = "." & formato
	
	'NOME DO ARQUIVO
	Filename = arquivo & "_" & data & extensao
	Excel = server.mappath("excel/" & Filename)

	set fso = createobject("scripting.filesystemobject")
	Set act = fso.CreateTextFile(Excel, true)
	act.WriteLine("<table cellpadding='2' cellspacing='1' border='1' style='font-family:Verdana, Geneva, sans-serif; font-size:12px;'>")
	act.WriteLine("<tr>")
	act.WriteLine("<td>Evento</td>")
	act.WriteLine("<td>Tipo</td>")
	act.WriteLine("<td>Ano</td>")
	act.WriteLine("<td>Nome</td>")
	act.WriteLine("<td>Curso</td>")
	act.WriteLine("<td>Email</td>")
	act.WriteLine("<td>Ativo</td>")
	act.WriteLine("<td>ID_Empresa</td>")
	act.WriteLine("<td>Razão Social</td>")
	act.WriteLine("<td>Nome Fantasia</td>")
	act.WriteLine("<td>Reitor</td>")
	act.WriteLine("<td>Nome Completo</td>")
	act.WriteLine("<td>Cargo</td>")
	act.WriteLine("</tr>")

	While Not RS_Alunos.Eof
	
	act.WriteLine("<tr>")
	act.WriteLine("<td>" &RS_Alunos("Evento")& "</td>")
	act.WriteLine("<td>" &RS_Alunos("Tipo")& "</td>")
	act.WriteLine("<td>" &RS_Alunos("Ano")& "</td>")
	act.WriteLine("<td>" &RS_Alunos("Nome")& "</td>")
	act.WriteLine("<td>" &RS_Alunos("Curso")& "</td>")
	act.WriteLine("<td>" &RS_Alunos("Email")& "</td>")
	act.WriteLine("<td>" &RS_Alunos("Ativo")& "</td>")
	act.WriteLine("<td>" &RS_Alunos("ID_Empresa")& "</td>")
	act.WriteLine("<td>" &RS_Alunos("Razao_Social")& "</td>")
	act.WriteLine("<td>" &RS_Alunos("Nome_Fantasia")& "</td>")
	act.WriteLine("<td>" &RS_Alunos("Reitor")& "</td>")
	act.WriteLine("<td>" &RS_Alunos("Nome_Completo")& "</td>")
	act.WriteLine("<td>" &RS_Alunos("Cargo_PTB")& "</td>")
	act.WriteLine("</tr>")
	
	RS_Alunos.MoveNext
	Wend
	
	act.WriteLine("</table>")
	act.close

	Response.Write(Filename)
Else
	Response.Write("ERRO")
End If
RS_Alunos.Close
%>

