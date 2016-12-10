<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Administra&ccedil;&atilde;o Cred. 2012</title>
</head>
<body  style="font-family:Arial, Helvetica, sans-serif; font-size:12px; ">
<%
DIM StartTime
Dim EndTime
StartTime = Timer()
%>
<%
Server.ScriptTimeout = 999999999

Function Zeros_ESQ (qtos, valor)
	For i = Len(valor) + 1 To qtos
		valor = "0" & valor
	Next
	Zeros_ESQ = valor
End Function 


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
data_inicio = d & "-" & m & "-" & a & "_" & h & "-" & n & "-" & s

id_edicao	= Request("id")

If IsNumeric(id_evento) = false Then response.Redirect("default.asp?msg=erro_nao_encontrado")
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================

	SQL_Evento_Autorizado = "Select " &_
							"	EE.Ano " &_
							"	,E.Nome_PTB as Feira " &_
							"From  Eventos_Edicoes as EE " &_
							"Inner Join Eventos as E ON E.ID_Evento = EE.ID_Evento " &_
							"Where  " &_
							"	Ee.ID_Edicao = " & id_edicao
	Set RS_Evento_Autorizado = Server.CreateObject("ADODB.Recordset")
	RS_Evento_Autorizado.Open SQL_Evento_Autorizado, Conexao
	
	If RS_Evento_Autorizado.BOF or RS_Evento_Autorizado.EOF Then
		response.Redirect("default.asp?msg=erro_nao_autorizado")
	Else
		Feira = Replace(RS_Evento_Autorizado("Ano") & "-" & RS_Evento_Autorizado("Feira"), " ", "_")
		Feira = Replace(Feira, "&", "")
		RS_Evento_Autorizado.Close
	End If

	%>
    <h1><%=feira%></h1><hr>
	<%
	response.Flush()

	SQL_Criar_Tabela_Temporaria = 	" " &_
									"IF (Not EXISTS (SELECT *  " &_
									"				 FROM INFORMATION_SCHEMA.TABLES  " &_
									"				 WHERE TABLE_SCHEMA = 'dbo'  " &_
									"				 AND  TABLE_NAME = 'Tabela_Temporaria_ContagemCPFs')) " &_
									"	BEGIN " &_
									"		SET ANSI_NULLS ON " &_
									"		SET QUOTED_IDENTIFIER ON " &_
									"		SET ANSI_PADDING ON " &_
									"		CREATE TABLE [dbo].[Tabela_Temporaria_ContagemCPFs]( " &_
									"			[ID] [numeric](18, 0) IDENTITY(1,1) NOT NULL, " &_
									"			[id_edicao] [numeric](18, 0) NULL, " &_
									"			[cpf] [varchar](20) NULL, " &_
									"			[registros] [numeric](18, 0) NULL, " &_
									"			[dt_inclusao] [datetime] NOT NULL " &_
									"		) " &_
									"		SET ANSI_PADDING OFF " &_
									"		ALTER TABLE [dbo].[Tabela_Temporaria_ContagemCPFs] ADD  CONSTRAINT [DF_Tabela_Temporaria_ContagemCPFs_dt_inclusao]  DEFAULT (getdate()) FOR [dt_inclusao] " &_
									"		ALTER TABLE [dbo].[Tabela_Temporaria_ContagemCPFs] ADD  CONSTRAINT [DF_Tabela_Temporaria_ContagemCPFs_registros]  DEFAULT (0) FOR [registros] " &_
									"	END " &_
									"ElSE " &_
									"	BEGIN " &_
									"		DELETE FROM [dbo].[Tabela_Temporaria_ContagemCPFs] WHERE id_edicao = " & id_edicao & " " &_
									"	END "

'	response.write(SQL_Criar_Tabela_Temporaria)

	Set RS_Criar_Tabela_Temporaria = Server.CreateObject("ADODB.Recordset")
		RS_Criar_Tabela_Temporaria.CursorType = 0
		RS_Criar_Tabela_Temporaria.LockType = 1
	RS_Criar_Tabela_Temporaria.Open SQL_Criar_Tabela_Temporaria, Conexao

	SQL_1_RelacionamentoCadastros =		"Select " &_
										"	* " &_
										"From Vw_Relacionamento_Cadastros_CPFCount " &_
										"Where " &_
										"	ID_Edicao = " & id_edicao  & " " &_
										"	AND Exportado in ('true', 'false')"

	Set RS_1_RelacionamentoCadastros = Server.CreateObject("ADODB.RecordSet")
	RS_1_RelacionamentoCadastros.CursorLocation = 3
	RS_1_RelacionamentoCadastros.CursorType = 3
	RS_1_RelacionamentoCadastros.LockType = 1
	RS_1_RelacionamentoCadastros.Open SQL_1_RelacionamentoCadastros, Conexao
	
	registros = 0

	Redim IDs(registros)
	If not RS_1_RelacionamentoCadastros.BOF or not RS_1_RelacionamentoCadastros.EOF Then
		i = 0
		While not RS_1_RelacionamentoCadastros.EOF
			i = i + 1
			RS_1_RelacionamentoCadastros.MoveNext
		Wend
		Redim IDs(i)
		RS_1_RelacionamentoCadastros.MoveFirst

		i = 0
		While not RS_1_RelacionamentoCadastros.EOF
			id 			= RS_1_RelacionamentoCadastros.Fields("ID_Relacionamento_Cadastro").Value
			idioma 		= RS_1_RelacionamentoCadastros.Fields("ID_Idioma").Value
			tipo		= RS_1_RelacionamentoCadastros.Fields("ID_Tipo_Credenciamento").Value
			visitante	= RS_1_RelacionamentoCadastros.Fields("ID_Visitante").Value
			empresa		= RS_1_RelacionamentoCadastros.Fields("ID_Empresa").Value
			formulario	= RS_1_RelacionamentoCadastros.Fields("ID_Formulario").Value
			data		= RS_1_RelacionamentoCadastros.Fields("Data_Cadastro").Value
			IDs(i) = Array(id, idioma, tipo, visitante, empresa, formulario, data)
			i = i + 1
			RS_1_RelacionamentoCadastros.MoveNext
		Wend
		RS_1_RelacionamentoCadastros.Close
		Set RS_1_RelacionamentoCadastros = Nothing
	End If
	%>
	<b>&bull; Registros <big><%=Ubound(IDs)%></big> listados...</b><br>
	<b>&bull; Buscando dados dos registros acima !</b><br>
	<% Intermediaria1 = Timer() %>
	<br><span style=''>Tempo da 1a Listagem: <%=FormatNumber((Intermediaria1 - StartTime),2)%> segundos</span><br>
	<div style="overflow:auto; width:850px; height:300px; border:1px solid #666; padding:5px; font-family:Arial, Helvetica, sans-serif; font-size:12px; " id="conteudo">
	<%	
	qtos_zeros = Len(Ubound(IDs))
	response.Flush() 
'			StartTime = Timer()

	total = 0
	'======================================================
	For x = Lbound(IDs) To Ubound(IDs) - 1
	
		'======================================================
		SQL_2_Visitantes = 	"Select " &_
							"	ID_Visitante, CPF, Nome_Completo " &_
							"From Vw_Visitantes " &_
							"Where  " &_
							"	ID_Visitante = " & IDs(x)(3)
'response.write("<hr><b>SQL_2_Visitantes</b><br>" & SQL_2_Visitantes & "<br>")
			
		Set RS_2_Visitantes = Server.CreateObject("ADODB.RecordSet")
		RS_2_Visitantes.CursorLocation = 3
		RS_2_Visitantes.CursorType = 3
		RS_2_Visitantes.LockType = 1
		RS_2_Visitantes.Open SQL_2_Visitantes, Conexao
		'======================================================
		
		'======================================================
		If not RS_2_Visitantes.BOF or not RS_2_Visitantes.EOF Then
			'		  Array(0,  1, 		2, 	  3, 		4,		  5,		  6)
			'IDs(i) = Array(id, idioma, tipo, visitante, empresa, formulario, data)
			ID_Cadastro		= IDs(x)(0)
			tipo_cadastro	= IDs(x)(2)
			id_empresa		= IDs(x)(4)
			formulario		= IDs(x)(5)
			data_cadastro	= IDs(x)(6)
			ID_Visitante 	= RS_2_Visitantes.Fields("ID_Visitante").Value
			CPF				= RS_2_Visitantes.Fields("CPF").Value
			Nome_Completo	= RS_2_Visitantes.Fields("Nome_Completo").Value
			
			SQL_Verificar_CPF = 	"Select " &_
									"	CPF " &_
									"From Tabela_Temporaria_ContagemCPFs " &_
									"Where " &_
									"	CPF = '"  & CPF & "' " &_
									"	AND id_edicao = " & id_Edicao
			Set RS_Verificar_CPF = Server.CreateObject("ADODB.RecordSet")
			RS_Verificar_CPF.CursorLocation = 3
			RS_Verificar_CPF.CursorType = 3
			RS_Verificar_CPF.LockType = 1
			RS_Verificar_CPF.Open SQL_Verificar_CPF, Conexao
			
			' Se não tiver GRAVE
			If RS_Verificar_CPF.BOF or RS_Verificar_CPF.EOF Then
				SQL_Inserir = 	"Insert into Tabela_Temporaria_ContagemCPFs " &_
								"(cpf, registros, id_edicao) " &_
								"Values " &_
								"('" & cpf & "', 1, " & id_edicao & ")" 
				Conexao.Execute(SQL_Inserir)
			Else
				SQL_Update = 	"Update Tabela_Temporaria_ContagemCPFs " &_
								"Set	registros = registros + 1 " &_
								"Where	" &_
								"	cpf = '" & cpf & "' " &_
								"	and id_edicao = " & id_edicao
				Conexao.Execute(SQL_Update)
				RS_Verificar_CPF.Close
				Set RS_Verificar_CPF = Nothing
			End If

			RS_2_Visitantes.Close
			Set RS_2_Visitantes = Nothing
		End If
		'======================================================
		%>
		<%=Zeros_ESQ(qtos_zeros,x+1)%> - <b>IDC:</b> <%=id_cadastro%> / <b>IDV:</b> <%=id_visitante%> / <b>CPF:</b> <%=CPF%> / <b>Nome:</b> <%=nome_completo%> <br>
		<script language="javascript">document.getElementById('conteudo').scrollTop += 100;</script>
		<%
		response.Flush() 
	Next
	'======================================================
	
	%>      
	</div>
	<%
	
	SQL_Contagem_CPFs = 	"Select " &_
							"	Sum(registros) as registros " &_
							"	,Count(CPF) as cpfs " &_
							"From Tabela_Temporaria_ContagemCPFs " &_
							"Where id_edicao = " & id_Edicao
	Set RS_Contagem_CPFs = Server.CreateObject("ADODB.RecordSet")
	RS_Contagem_CPFs.CursorLocation = 3
	RS_Contagem_CPFs.CursorType = 3
	RS_Contagem_CPFs.LockType = 1
	RS_Contagem_CPFs.Open SQL_Contagem_CPFs, Conexao
	
	%>
<hr>
<table width="400" border="1" cellspacing="0" cellpadding="10" style="font-family:Arial, Helvetica, sans-serif; font-size:12px; ">
  <tr>
    <td>Registros LOOP (banco de dados)</td>
    <td width="80" align="right"><%=Ubound(IDs)%></td>
  </tr>
  <tr>
    <td>Total de registros tabela TEMP</td>
    <td width="80" align="right"><%=RS_Contagem_CPFs.Fields("registros").Value%></td>
  </tr>
  <tr>
    <td>Total de CPFS</td>
    <td width="80" align="right"><%=RS_Contagem_CPFs.Fields("cpfs").Value%></td>
  </tr>
</table>
	<% 
	RS_Contagem_CPFs.Close
	EndTime = Timer()
	response.write("<br><br><span style='padding-left:50px;'>Tempo de processamento: " & FormatNumber((EndTime - StartTime),2) & " segundos</span>") 
	%>
</body>
</html>

<%
Conexao.Close
Set Conexao = Nothing
%>