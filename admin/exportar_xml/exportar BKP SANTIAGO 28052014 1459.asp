<%
DIM StartTime
Dim EndTime
StartTime = Timer()
%>
<!--#include virtual="/admin/inc/logado.asp"-->
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
Server.ScriptTimeout = 999999999

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

id_edicao	= Limpar_Texto(Request("id"))

Function RemoverAcentuacao (texto)
	limpar = Ucase(texto)
	If Len(limpar) <= 0 or Isnull(limpar) Then
	Else
		limpar = Replace(limpar, "�", "A")
		limpar = Replace(limpar, "�", "A")
		limpar = Replace(limpar, "�", "A")
		limpar = Replace(limpar, "�", "A")
		limpar = Replace(limpar, "�", "A")
		limpar = Replace(limpar, "�", "E")
		limpar = Replace(limpar, "�", "E")
		limpar = Replace(limpar, "�", "E")
		limpar = Replace(limpar, "�", "I")
		limpar = Replace(limpar, "�", "I")
		limpar = Replace(limpar, "�", "I")
		limpar = Replace(limpar, "�", "O")
		limpar = Replace(limpar, "�", "O")
		limpar = Replace(limpar, "�", "O")
		limpar = Replace(limpar, "�", "O")
		limpar = Replace(limpar, "�", "U")
		limpar = Replace(limpar, "�", "U")
		limpar = Replace(limpar, "�", "U")
		limpar = Replace(limpar, "�", "C")
		limpar = Replace(limpar, "&", "E")
	End If
	RemoverAcentuacao = limpar
End Function

Function trocar(qual)
	If Len(qual) > 0 Then
		limpar = qual 
		limpar = Replace(limpar, "&", "&amp;")
		limpar = Replace(limpar, "<", "&lt;")
		limpar = Replace(limpar, ">", "&gt;")
		limpar = Replace(limpar, """", "&quot;")
		trocar = limpar
	Else
		trocar = qual
	End If
End Function

Function limpar_formatacao(qual)
	If Len(qual) > 0 Then
		limpar = qual 
		limpar = Replace(limpar, ".", "")
		limpar = Replace(limpar, "-", "")
		limpar = Replace(limpar, "/", "")
		limpar_formatacao = limpar
	Else
		limpar_formatacao = qual
	End If	
End Function

Function Zeros_ESQ (qtos, valor)
	For i = Len(valor) + 1 To qtos
		valor = "0" & valor
	Next
	Zeros_ESQ = valor
End Function 

If IsNumeric(id_evento) = false Then response.Redirect("default.asp?msg=erro_nao_encontrado")
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================

	SQL_Evento_Autorizado = "Select " &_
							"	EE.Ano " &_
							"	,E.Nome_PTB as Feira " &_
							"From Administradores_Edicoes AE " &_
							"Inner Join Eventos_Edicoes as EE ON Ae.ID_Edicao = EE.ID_Edicao " &_
							"Inner Join Eventos as E ON E.ID_Evento = EE.ID_Evento " &_
							"Where  " &_
							"	Ae.ID_Admin = " & session("admin_id_usuario") & " " &_
							"	AND Ae.ID_Edicao = " & id_edicao
	Set RS_Evento_Autorizado = Server.CreateObject("ADODB.Recordset")
	RS_Evento_Autorizado.Open SQL_Evento_Autorizado, Conexao
	
	If RS_Evento_Autorizado.BOF or RS_Evento_Autorizado.EOF Then
		response.Redirect("default.asp?msg=erro_nao_autorizado")
	Else
		Feira = Replace(RS_Evento_Autorizado("Ano") & "-" & RS_Evento_Autorizado("Feira"), " ", "_")
		sAnoPasta=RS_Evento_Autorizado("Ano")
		Feira = Replace(Feira, "&", "")
		RS_Evento_Autorizado.Close
	End If

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Administra&ccedil;&atilde;o Cred. 2012</title>
</head>
<link href="/admin/css/bts.css" rel="stylesheet" type="text/css" />
<link href="/admin/css/admin.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
<script language="javascript">
$(document).ready(function(){
	$('#aviso').hide();
});
</script>

<body>
<!--#include virtual="/admin/inc/menu_top.asp"-->
<table width="955" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><img src="/admin/images/img_tabela_branca_top.jpg" width="955" height="15" /></td>
  </tr>
</table>
<table width="955" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td align="left" bgcolor="#FFFFFF" class="conteudo_site" style="padding:20px;">
    
    <table width="900" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="100" height="50">&nbsp;</td>
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">Exportar XML<br>
          Evento: <%=Feira%></span></td>
        <td width="100" align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;"><a href="evento.asp?id=<%=id_edicao%>"><img src="/admin/images/bt_voltar_admin.gif" width="100" height="48" border="0" /></a></td>
      </tr>
    </table>
      <div id="aviso" style="background-color:#FFFF00; width:100%; text-align:center;" class="t_arial fs11px bold c_preto"> <img src="/admin/images/ico_aviso.gif" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso"> <span id="aviso_conteudo">Aviso Processando, por favor n�o interrompa.</span></div>

		<%
		response.Flush()

		SQL_1_RelacionamentoCadastros =		"Select * " &_
											"From Vw_Relacionamento_Cadastros " &_
											"Where " &_
											"	ID_Edicao = " & id_edicao 
											
'		SQL_1_RelacionamentoCadastros =	"SELECT	 RC.ID_Relacionamento_Cadastro" & VBCRLF & _
'										"		,RC.ID_Idioma" & VBCRLF & _
'										"		,RC.ID_Tipo_Credenciamento" & VBCRLF & _
'										"		,RC.ID_Visitante" & VBCRLF & _
'										"		,RC.ID_Empresa" & VBCRLF & _
'										"		,RC.ID_Edicao" & VBCRLF & _
'										"		,TC.ID_Formulario" & VBCRLF & _
'										"		,RC.Data_Cadastro" & VBCRLF & _
'										"		,RC.Exportado" & VBCRLF & _
'										"FROM dbo.Relacionamento_Cadastro AS RC " & VBCRLF & _
'										"INNER JOIN dbo.Tipo_Credenciamento AS TC " & VBCRLF & _
'										"	ON TC.ID_Tipo_Credenciamento = RC.ID_Tipo_Credenciamento" & VBCRLF & _
'										"WHERE ID_Empresa='85531' AND ID_Edicao = " & id_edicao


		Set RS_1_RelacionamentoCadastros = Server.CreateObject("ADODB.RecordSet")
		RS_1_RelacionamentoCadastros.CursorLocation = 2
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
		<br><span style=''>Tempo da 1� Listagem: <%=FormatNumber((Intermediaria1 - StartTime),2)%> segundos</span><br>
        <div style="overflow:auto; width:850px; height:300px; border:1px solid #666; padding:5px;" id="conteudo">
        <%	
			qtos_zeros = Len(Ubound(IDs))
			response.Flush() 
'			StartTime = Timer()
		%>
        <%
		
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
	
		arquivo = "pre_credenciados"
						
		extensao = ".xml"
	
		Filename = arquivo & "_" & RemoverAcentuacao(feira) & "_" & data & extensao ' file to read
	
		Const ForReading = 1, ForWriting = 2, ForAppending = 3
		Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
	
		' Create a filesystem object
		Dim FSO
		set FSO = server.createObject("Scripting.FileSystemObject")
	
		' Map the logical path to the physical system path
		Dim Filepath


		'Filepath = Server.MapPath("arquivos_2012/" & Filename)
		Filepath = Server.MapPath("arquivos_" & sAnoPasta & "/" & Filename)
	
	
		Set oFiletxt = FSO.CreateTextFile(Filepath, True)
		sPath = FSO.GetAbsolutePathName(Filepath)
		sFilename = FSO.GetFileName(sPath)
	
		oFiletxt.WriteLine("<?xml version='1.0' encoding='iso-8859-1'?>")
		oFiletxt.WriteLine("<credenciamento>")
		total = 0
		'======================================================
		For x = Lbound(IDs) To Ubound(IDs) - 1
		
			'======================================================
			SQL_2_Visitantes = 	"Select " &_
								"	* " &_
								"From Vw_Visitantes " &_
								"Where  " &_
								"	ID_Visitante = " & IDs(x)(3)
'response.write("<hr><b>SQL_2_Visitantes</b><br>" & SQL_2_Visitantes & "<br>")
				
			Set RS_2_Visitantes = Server.CreateObject("ADODB.RecordSet")
			RS_2_Visitantes.CursorLocation = 2
			RS_2_Visitantes.CursorType = 3
			RS_2_Visitantes.LockType = 1
			RS_2_Visitantes.Open SQL_2_Visitantes, Conexao
			'======================================================

			'======================================================
			' LIMPANDO DADOS DE EMPRESA
			Empresa_Qtde_Funcionarios 	= ""
			Empresa_CNPJ				= ""
			Empresa_Razao_Social		= ""
			Empresa_Nome_Fantasia		= ""
			Empresa_Principal_Produto	= ""
			Empresa_Site  				= ""
			Empresa_Presidente			= ""
			Empresa_Reitor				= ""
			Empresa_Senha				= ""
			Empresa_Tipo_Telefone 	= ""
			Empresa_DDI 			= ""
			Empresa_DDD 			= ""
			Empresa_Numero 			= ""
			Empresa_Ramal 			= ""
			Empresa_SMS 			= ""
			Ramos					= ""
			Atividadades			= ""
			InteresseFeira			= ""
			Endereco_CEP 			= ""
			Endereco_Endereco 		= ""
			Endereco_Numero 		= ""
			Endereco_Complemento 	= ""
			Endereco_Bairro 		= ""
			Endereco_Cidade 		= ""
			Endereco_UF				= ""
			Endereco_Pais			= ""
			
			' Regra sobre TIPO_FORMULARIO
			' Se for diferente de 4 busque TELEFONES da EMPRESA e DADOS da EMPRESA
			If IDs(x)(5) <> "4" AND Len(Trim(IDs(x)(4))) > 0 Then
				'======================================================
				SQL_3_Telefones = 	"Select " &_
									"	* " &_
									"From Vw_Telefones " &_
									"Where  " &_
									"	ID_Empresa = " & IDs(x)(4)
					
				Set RS_3_Telefones = Server.CreateObject("ADODB.RecordSet")
				RS_3_Telefones.CursorLocation = 2
				RS_3_Telefones.CursorType = 3
				RS_3_Telefones.LockType = 1
				RS_3_Telefones.Open SQL_3_Telefones, Conexao
				'======================================================
				
				'======================================================
				If not RS_3_Telefones.BOF or not RS_3_Telefones.EOF Then
					Empresa_Tipo_Telefone 	= RS_3_Telefones.Fields("Tipo_Telefone").Value
					Empresa_DDI 			= RS_3_Telefones.Fields("DDI").Value
					Empresa_DDD 			= RS_3_Telefones.Fields("DDD").Value
					Empresa_Numero 			= RS_3_Telefones.Fields("Numero").Value
					Empresa_Ramal 			= RS_3_Telefones.Fields("Ramal").Value
					Empresa_SMS 			= RS_3_Telefones.Fields("SMS").Value
					
					RS_3_Telefones.Close
					Set RS_3_Telefones = Nothing
				Else
					Empresa_Tipo_Telefone 	= ""
					Empresa_DDI 			= ""
					Empresa_DDD 			= ""
					Empresa_Numero 			= ""
					Empresa_Ramal 			= ""
					Empresa_SMS 			= ""
				End If
				'======================================================
				
				'======================================================
				' EMPRESAS
				'		  Array(0,  1, 		2, 	  3, 		4,		  5,		  6)
				'IDs(i) = Array(id, idioma, tipo, visitante, empresa, formulario, data)
				'======================================================
				SQL_6_Empresa = 	"Select " &_
									"	* " &_
									"From vw_empresas " &_
									"Where  " &_
									"	ID_Empresa = " & IDs(x)(4)
					
				Set RS_6_Empresa = Server.CreateObject("ADODB.RecordSet")
				RS_6_Empresa.CursorLocation = 2
				RS_6_Empresa.CursorType = 3
				RS_6_Empresa.LockType = 1
				RS_6_Empresa.Open SQL_6_Empresa, Conexao
				'======================================================
				
				'======================================================
				If not RS_6_Empresa.BOF or not RS_6_Empresa.EOF Then
					Empresa_Qtde_Funcionarios 	= RS_6_Empresa.Fields("Qtde_Funcionarios").Value
					Empresa_CNPJ				= RS_6_Empresa.Fields("CNPJ").Value
					Empresa_Razao_Social		= RS_6_Empresa.Fields("Razao_Social").Value
					Empresa_Nome_Fantasia		= RS_6_Empresa.Fields("Nome_Fantasia").Value
					Empresa_Principal_Produto	= "" 'RS_6_Empresa.Fields("Principal_Produto").Value
					Empresa_Site  				= RS_6_Empresa.Fields("Site").Value
					Empresa_Presidente			= RS_6_Empresa.Fields("Presidente").Value
					Empresa_Reitor				= RS_6_Empresa.Fields("Reitor").Value
					Empresa_Senha				= RS_6_Empresa.Fields("Senha").Value
					
					RS_6_Empresa.Close
					Set RS_6_Empresa = Nothing
				Else
					Empresa_Qtde_Funcionarios 	= ""
					Empresa_CNPJ				= ""
					Empresa_Razao_Social		= ""
					Empresa_Nome_Fantasia		= ""
					Empresa_Principal_Produto	= ""
					Empresa_Site  				= ""
					Empresa_Presidente			= ""
					Empresa_Reitor				= ""
					Empresa_Senha				= ""
				End If
				'======================================================
				
				'======================================================
				SQL_7_Empresa_Produtos = 	"Select " &_
											"	* " &_
											"From Vw_Principal_Produto_2 " &_
											"Where  ID_Empresa = " & IDs(x)(4) &_
											"AND ID_Edicao=" & id_edicao
					
				Set RS_7_Empresa_Produtos = Server.CreateObject("ADODB.RecordSet")
				RS_7_Empresa_Produtos.CursorLocation = 2
				RS_7_Empresa_Produtos.CursorType = 3
				RS_7_Empresa_Produtos.LockType = 1
				RS_7_Empresa_Produtos.Open SQL_7_Empresa_Produtos, Conexao
				'======================================================
				
				'======================================================
				Empresa_Principal_Produto = ""
				If not RS_7_Empresa_Produtos.BOF or not RS_7_Empresa_Produtos.EOF Then
					While not RS_7_Empresa_Produtos.EOF
						Empresa_Principal_Produto = Empresa_Principal_Produto & RS_7_Empresa_Produtos.Fields("Principal_Produto").Value
						RS_7_Empresa_Produtos.MoveNext
						If not RS_7_Empresa_Produtos.EOF Then Empresa_Principal_Produto = Empresa_Principal_Produto & "; "
					Wend
					RS_7_Empresa_Produtos.Close
					Set RS_7_Empresa_Produtos = Nothing
				End If
				'======================================================
				
				'======================================================
				SQL_8_Empresa_Ramos = 	"Select " &_
										"	* " &_
										"From [Vw_Ramo-Atividade_2] " &_
										"Where  " &_
										"	ID_Empresa = " & IDs(x)(4) & " " &_
										"	AND ID_Edicao = " & id_edicao & " " &_
										"Order by Ramo_Atv_PTB"
				'response.Write("<br>"&SQL_8_Empresa_Ramos&"<br>")
				'response.End()
				
					
				Set RS_8_Empresa_Ramos = Server.CreateObject("ADODB.RecordSet")
				RS_8_Empresa_Ramos.CursorLocation = 3
				RS_8_Empresa_Ramos.CursorType = 3
				RS_8_Empresa_Ramos.LockType = 1
				RS_8_Empresa_Ramos.Open SQL_8_Empresa_Ramos, Conexao
				'======================================================
				
				'======================================================
				qtde_ramos = 0
				If not RS_8_Empresa_Ramos.BOF or not RS_8_Empresa_Ramos.EOF Then
					While not RS_8_Empresa_Ramos.EOF
						qtde_ramos = qtde_ramos + 1
						RS_8_Empresa_Ramos.MoveNext
					Wend
					RS_8_Empresa_Ramos.MoveFirst
			
					Redim Ramos_Atividades(qtde_ramos)

					i = 0
					While not RS_8_Empresa_Ramos.EOF
						Ramo			= RS_8_Empresa_Ramos.Fields("Ramo_Atv_PTB").Value
						Ramo_Outro		= RS_8_Empresa_Ramos.Fields("Complemento").Value
							Ramos_Atividades(i) = Array(Ramo & " - " & Ramo_Outro)
							i = i + 1
						RS_8_Empresa_Ramos.MoveNext
					Wend
					RS_8_Empresa_Ramos.Close
					Set RS_8_Empresa_Ramos = Nothing
				End If
				'======================================================
				
				'======================================================
				SQL_10_Empresa_InteresseFeira = 	"Select " &_
													"	* " &_
													"From Vw_InteresseFeira " &_
													"Where  " &_
													"	ID_Relacionamento_Cadastro = " & IDs(x)(0)
					
				Set RS_10_Empresa_InteresseFeira = Server.CreateObject("ADODB.RecordSet")
				RS_10_Empresa_InteresseFeira.CursorLocation = 2
				RS_10_Empresa_InteresseFeira.CursorType = 3
				RS_10_Empresa_InteresseFeira.LockType = 1
				RS_10_Empresa_InteresseFeira.Open SQL_10_Empresa_InteresseFeira, Conexao
				'======================================================
				
				'======================================================
				InteresseFeira = ""
				If not RS_10_Empresa_InteresseFeira.BOF or not RS_10_Empresa_InteresseFeira.EOF Then
					While not RS_10_Empresa_InteresseFeira.EOF
						InteresseFeira 		= InteresseFeira & RS_10_Empresa_InteresseFeira.Fields("AreaInteresse").Value
						RS_10_Empresa_InteresseFeira.MoveNext
						If not RS_10_Empresa_InteresseFeira.EOF Then InteresseFeira = InteresseFeira & "; "
					Wend
					RS_10_Empresa_InteresseFeira.Close
					Set RS_10_Empresa_InteresseFeira = Nothing
				End If
				'======================================================
						
			End If
			'======================================================	

			'======================================================
			SQL_11_Perguntas = 	"Select " &_
								"	* " &_
								"From Vw_Respostas_Perguntas " &_
								"Where  " &_
								"	ID_Relacionamento_Cadastro = " & IDs(x)(0)
			
			'response.write(SQL_11_Perguntas & "<br>")
				
			Set RS_11_Perguntas = Server.CreateObject("ADODB.RecordSet")
			RS_11_Perguntas.CursorLocation = 2
			RS_11_Perguntas.CursorType = 3
			RS_11_Perguntas.LockType = 1
			RS_11_Perguntas.Open SQL_11_Perguntas, Conexao
			'======================================================
			
			'======================================================
			If not RS_11_Perguntas.BOF or not RS_11_Perguntas.EOF Then
				Pergunta_OLD = ""
				qtde_perguntas = 0
				While not RS_11_Perguntas.EOF
					Pergunta_Atual 	= RS_11_Perguntas.Fields("Pergunta").Value
					If Trim(Pergunta_OLD) <> Trim(Pergunta_Atual) Then 
						Pergunta_OLD = Pergunta_Atual
						qtde_perguntas = qtde_perguntas + 1
					End If
					RS_11_Perguntas.MoveNext
				Wend
				RS_11_Perguntas.MoveFirst
		
				Redim Perguntas_e_Respostas(qtde_perguntas)
				Pergunta_OLD = ""
				Todas_Respostas = ""
				i = 0
				While not RS_11_Perguntas.EOF
					Pergunta_Atual 	= RS_11_Perguntas.Fields("Pergunta").Value
					Resposta		= RS_11_Perguntas.Fields("Resposta").Value
		
					If Trim(Pergunta_OLD) <> Trim(Pergunta_Atual) Then 
						Pergunta_OLD = Pergunta_Atual
						i = i + 1
						Todas_Respostas = ""
						Todas_Respostas = Todas_Respostas & Resposta
						Perguntas_e_Respostas(i) = Array(Pergunta_Atual, Todas_Respostas)
					ElseIf Trim(Pergunta_OLD) = Trim(Pergunta_Atual) Then 
						Todas_Respostas = Todas_Respostas & "; " & Resposta
						Perguntas_e_Respostas(i) = Array(Pergunta_Atual, Todas_Respostas)
					End If
					RS_11_Perguntas.MoveNext
				Wend
				RS_11_Perguntas.Close
				Set RS_11_Perguntas = Nothing
			
			Else 
				'======================================================
				SQL_11_Perguntas = 	"Select " &_
									"	* " &_
									"From Vw_Respostas_Perguntas " &_
									"Where  " &_
									"	ID_Relacionamento_Cadastro = " & IDs(x)(3)
				
				response.write(SQL_11_Perguntas & "<br>")
					
				Set RS_11_Perguntas = Server.CreateObject("ADODB.RecordSet")
				RS_11_Perguntas.CursorLocation = 2
				RS_11_Perguntas.CursorType = 3
				RS_11_Perguntas.LockType = 1
				RS_11_Perguntas.Open SQL_11_Perguntas, Conexao
				'======================================================

				'======================================================
					If not RS_11_Perguntas.BOF or not RS_11_Perguntas.EOF Then
						Pergunta_OLD = ""
						qtde_perguntas = 0
						While not RS_11_Perguntas.EOF
							Pergunta_Atual 	= RS_11_Perguntas.Fields("Pergunta").Value
							If Trim(Pergunta_OLD) <> Trim(Pergunta_Atual) Then 
								Pergunta_OLD = Pergunta_Atual
								qtde_perguntas = qtde_perguntas + 1
							End If
							RS_11_Perguntas.MoveNext
						Wend
						RS_11_Perguntas.MoveFirst
				
						Redim Perguntas_e_Respostas(qtde_perguntas)
						Pergunta_OLD = ""
						Todas_Respostas = ""
						i = 0
						While not RS_11_Perguntas.EOF
							Pergunta_Atual 	= RS_11_Perguntas.Fields("Pergunta").Value
							Resposta		= RS_11_Perguntas.Fields("Resposta").Value
				
							If Trim(Pergunta_OLD) <> Trim(Pergunta_Atual) Then 
								Pergunta_OLD = Pergunta_Atual
								i = i + 1
								Todas_Respostas = ""
								Todas_Respostas = Todas_Respostas & Resposta
								Perguntas_e_Respostas(i) = Array(Pergunta_Atual, Todas_Respostas)
							ElseIf Trim(Pergunta_OLD) = Trim(Pergunta_Atual) Then 
								Todas_Respostas = Todas_Respostas & "; " & Resposta
								Perguntas_e_Respostas(i) = Array(Pergunta_Atual, Todas_Respostas)
							End If
							RS_11_Perguntas.MoveNext
						Wend
						RS_11_Perguntas.Close
						Set RS_11_Perguntas = Nothing
					End If
			End If
			'======================================================

			'======================================================
			' Regra sobre TIPO_FORMULARIO
			' Se for diferente de 4 BUSQUE ENDERECO da EMPRESA se NAO pegue do VISITANTE
			'		  Array(0,  1, 		2, 	  3, 		4,		  5,		  6)
			'IDs(i) = Array(id, idioma, tipo, visitante, empresa, formulario, data)
			
			If IDs(x)(5) <> "4" AND Len(Trim(IDs(x)(4))) > 0 Then
				Where_Endereco = "	ID_Empresa = " & IDs(x)(4)
			Else
				Where_Endereco = "	ID_Visitante = " & IDs(x)(3)
			End If
			'======================================================
				
			'======================================================
			SQL_5_Endereco = 	"Select " &_
								"	* " &_
								"From Vw_Enderecos " &_
								"Where  " &_
								" " & Where_Endereco
				
			Set RS_5_Endereco = Server.CreateObject("ADODB.RecordSet")
			RS_5_Endereco.CursorLocation = 2
			RS_5_Endereco.CursorType = 3
			RS_5_Endereco.LockType = 1
			RS_5_Endereco.Open SQL_5_Endereco, Conexao
			'======================================================
			
			'======================================================
			If not RS_5_Endereco.BOF or not RS_5_Endereco.EOF Then
				Endereco_CEP			= RS_5_Endereco.Fields("CEP").Value
				Endereco_Endereco 		= RS_5_Endereco.Fields("Endereco").Value
				Endereco_Numero 		= RS_5_Endereco.Fields("Numero").Value
				Endereco_Complemento 	= RS_5_Endereco.Fields("Complemento").Value
				Endereco_Bairro 		= RS_5_Endereco.Fields("Bairro").Value
				Endereco_Cidade 		= RS_5_Endereco.Fields("Cidade").Value
				Endereco_UF				= RS_5_Endereco.Fields("UF").Value
				Endereco_Pais			= RS_5_Endereco.Fields("Pais").Value
				
				RS_5_Endereco.Close
				Set RS_5_Endereco = Nothing
			Else
				Endereco_CEP 			= ""
				Endereco_Endereco 		= ""
				Endereco_Numero 		= ""
				Endereco_Complemento 	= ""
				Endereco_Bairro 		= ""
				Endereco_Cidade 		= ""
				Endereco_UF				= ""
				Endereco_Pais			= ""
			End If
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
				Passaporte		= RS_2_Visitantes.Fields("Passaporte").Value
				Nome_Completo	= RS_2_Visitantes.Fields("Nome_Completo").Value
				Nome_Credencial	= RS_2_Visitantes.Fields("Nome_Credencial").Value
				Data_Nasc		= RS_2_Visitantes.Fields("Data_Nasc").Value
				Sexo			= RS_2_Visitantes.Fields("Sexo").Value
				
				If Sexo = "0" Then Sexo = "Masculino"
				If Sexo = "1" Then Sexo = "Feminino"
				
				Email			= RS_2_Visitantes.Fields("Email").Value
				Newsletter		= RS_2_Visitantes.Fields("Newsletter").Value
				
				If Newsletter = "0" Then Newsletter = "Nao"
				If Newsletter = "1" Then Newsletter = "Sim"
				
				Cargo			= RS_2_Visitantes.Fields("Cargo").Value
				Cargo_Outros	= RS_2_Visitantes.Fields("Cargo_Outros").Value
				SubCargo		= RS_2_Visitantes.Fields("SubCargo").Value
				SubCargo_Outros	= RS_2_Visitantes.Fields("SubCargo_Outros").Value
				Depto			= RS_2_Visitantes.Fields("Depto").Value
				Depto_Outros	= RS_2_Visitantes.Fields("Depto_Outros").Value
				
				'======================================================
				SQL_4_Telefones = 	"Select " &_
									"	* " &_
									"From Vw_Telefones " &_
									"Where  " &_
									"	ID_Visitante = " & IDs(x)(3)
					
				Set RS_4_Telefones = Server.CreateObject("ADODB.RecordSet")
				RS_4_Telefones.CursorLocation = 2
				RS_4_Telefones.CursorType = 3
				RS_4_Telefones.LockType = 1
				RS_4_Telefones.Open SQL_4_Telefones, Conexao
				'======================================================
				
				'======================================================
				If not RS_4_Telefones.BOF or not RS_4_Telefones.EOF Then
					i = 0
					While not RS_4_Telefones.EOF
						i = i + 1
						RS_4_Telefones.MoveNext
					Wend
					Redim Visitante_Telefones(i)
					RS_4_Telefones.MoveFirst

					i = 0
					While not RS_4_Telefones.EOF
						Visitante_Tipo_Telefone = RS_4_Telefones.Fields("Tipo_Telefone").Value
						Visitante_DDI 			= RS_4_Telefones.Fields("DDI").Value
						Visitante_DDD 			= RS_4_Telefones.Fields("DDD").Value
						Visitante_Numero		= RS_4_Telefones.Fields("Numero").Value
						Visitante_Ramal 		= RS_4_Telefones.Fields("Ramal").Value
						Visitante_SMS 			= RS_4_Telefones.Fields("SMS").Value
						
						Visitante_Telefones(i) = Array(Visitante_Tipo_Telefone, Visitante_DDI, Visitante_DDD, Visitante_Numero, Visitante_Ramal, Visitante_SMS)
						RS_4_Telefones.MoveNext
						i = i + 1
					Wend					
					RS_4_Telefones.Close
					Set RS_4_Telefones = Nothing
				End If
				'======================================================
				total = total + 1
							
				oFiletxt.WriteLine("<cadastro>")
					oFiletxt.WriteLine("<id_cadastro>" 			& trocar(Ucase( ID_Cadastro )) & "</id_cadastro>")
					oFiletxt.WriteLine("<id_visitante>" 		& trocar(Ucase( ID_Visitante )) & "</id_visitante>")
					oFiletxt.WriteLine("<id_empresa>" 			& trocar(Ucase( id_empresa )) & "</id_empresa>")
					oFiletxt.WriteLine("<formulario>" 			& trocar(Ucase( formulario )) & "</formulario>")
					oFiletxt.WriteLine("<data_cadastro>" 		& trocar(Ucase( data_cadastro )) & "</data_cadastro>")
					
					oFiletxt.WriteLine("<cnpj>" 				& limpar_formatacao(trocar(Ucase( Empresa_CNPJ ))) & "</cnpj>")
					oFiletxt.WriteLine("<razao_social>"			& trocar(Ucase( Empresa_Razao_Social )) & "</razao_social>")
					oFiletxt.WriteLine("<nome_fantasia>"		& trocar(Ucase( Empresa_Nome_Fantasia )) & "</nome_fantasia>")
					oFiletxt.WriteLine("<qtde_funcionarios>"	& trocar(Ucase( Empresa_Qtde_Funcionarios )) & "</qtde_funcionarios>")
					oFiletxt.WriteLine("<principal_produto>"	& trocar(Ucase( Empresa_Principal_Produto )) & "</principal_produto>")

'	MODELO ANTIGO					
'					oFiletxt.WriteLine("<ramos>"				& trocar(Ucase( Ramos )) & "</ramos>")
'					oFiletxt.WriteLine("<atividades>"			& trocar(Ucase( Atividadades )) & "</atividades>")


'	MODELO NOVO 12 07 2012 20:30 HOMERO
'					oFiletxt.WriteLine("<ramos_atividades>")
'					If qtde_ramos > 0 Then
'						For item_ramo = Lbound(Ramos_Atividades) To Ubound(Ramos_Atividades)-1
'							oFiletxt.WriteLine("<item ramo_atividade='" & RemoverAcentuacao(Ramos_Atividades(item_ramo)(0)) & "' ramo_outros='" & RemoverAcentuacao(Ramos_Atividades(item_ramo)(1)) & "' atividade_outros='" & RemoverAcentuacao(Ramos_Atividades(item_ramo)(2)) & "' />")
'						Next
'					End If
'					oFiletxt.WriteLine("</ramos_atividades>")
					
'	MODELO NOVO 20130408 WAGNER
					oFiletxt.WriteLine("<ramos_atividades>")
					If qtde_ramos > 0 Then
						w=0
						'response.Write("qtde_ramos="&qtde_ramos&"<br>")
						while w < Ubound(Ramos_Atividades)
							arrRamAtiv = split(RemoverAcentuacao(Ramos_Atividades(w)(0))," - ")
							oFiletxt.WriteLine("<item ramo_atividade='" & RemoverAcentuacao(arrRamAtiv(0)) & "' ramo_outros='" & RemoverAcentuacao(arrRamAtiv(1)) & "' atividade_outros='' />")
							'response.Write("ramo_atividade="&RemoverAcentuacao(arrRamAtiv(0))&"<br>")
							'response.Write("ramo_outros="&RemoverAcentuacao(arrRamAtiv(1))&"<br>")
							
							w=w+1
						wend
						'response.Write(w)
					End If
					oFiletxt.WriteLine("</ramos_atividades>")
					
					oFiletxt.WriteLine("<interesses_feira>"		& trocar(Ucase( InteresseFeira )) & "</interesses_feira>")
					oFiletxt.WriteLine("<site>"					& trocar(Ucase( Empresa_Site )) & "</site>")
					oFiletxt.WriteLine("<presidente>"			& trocar(Ucase( Empresa_Presidente )) & "</presidente>")
					oFiletxt.WriteLine("<reitor>"				& trocar(Ucase( Empresa_Reitor )) & "</reitor>")

					oFiletxt.WriteLine("<tipo_tel_1>" 			& trocar(Ucase( Empresa_Tipo_Telefone )) & "</tipo_tel_1>")
					oFiletxt.WriteLine("<ddi_1>" 				& trocar(Ucase( Empresa_DDI )) & "</ddi_1>")
					oFiletxt.WriteLine("<ddd_1>" 				& trocar(Ucase( Empresa_DDD )) & "</ddd_1>")
					oFiletxt.WriteLine("<fone_1>" 				& limpar_formatacao(trocar(Ucase( Empresa_Numero ))) & "</fone_1>")					
					oFiletxt.WriteLine("<ramal_1>" 				& trocar(Ucase( Empresa_Ramal )) & "</ramal_1>")
					oFiletxt.WriteLine("<sms_1>" 				& trocar(Ucase( Empresa_SMS )) & "</sms_1>")

				If idioma = "1" AND Len(Trim(CPF)) > 0 Then
					oFiletxt.WriteLine("<cpf>" 					& limpar_formatacao(trocar(Ucase( CPF ))) & "</cpf>")
				Else
					QtdeCPF = Len(ID_Visitante)
					RelInternacional = ID_Visitante
					
					For i=QtdeCPF to 10 
						RelInternacional = "0" & RelInternacional
					Next 
					
					oFiletxt.WriteLine("<cpf>" 					& trocar(Ucase( RelInternacional )) & "</cpf>")
				End If
					oFiletxt.WriteLine("<passaporte>"			& limpar_formatacao(trocar(Ucase( Passaporte ))) & "</passaporte>")
					oFiletxt.WriteLine("<nome>" 				& trocar(Ucase( Nome_Completo )) & "</nome>")
					oFiletxt.WriteLine("<credencial>" 			& trocar(Ucase( Nome_Credencial )) & "</credencial>")
					oFiletxt.WriteLine("<data_nasc>" 			& trocar(Ucase( Data_Nasc )) & "</data_nasc>")
					oFiletxt.WriteLine("<sexo>" 				& trocar(Ucase( Sexo )) & "</sexo>")
					oFiletxt.WriteLine("<email>" 				& trocar(Ucase( Email )) & "</email>")
					oFiletxt.WriteLine("<newsletter>" 			& trocar(Ucase( Newsletter )) & "</newsletter>")

					oFiletxt.WriteLine("<tipo_tel_2>" 			& trocar(Ucase( Visitante_Telefones(0)(0) )) & "</tipo_tel_2>")
					oFiletxt.WriteLine("<ddi_2>" 				& trocar(Ucase( Visitante_Telefones(0)(1) )) & "</ddi_2>")
					oFiletxt.WriteLine("<ddd_2>" 				& trocar(Ucase( Visitante_Telefones(0)(2) )) & "</ddd_2>")
					oFiletxt.WriteLine("<fone_2>" 				& limpar_formatacao(trocar(Ucase( Visitante_Telefones(0)(3) ))) & "</fone_2>")					
					oFiletxt.WriteLine("<ramal_2>" 				& trocar(Ucase( Visitante_Telefones(0)(4) )) & "</ramal_2>")
					oFiletxt.WriteLine("<sms_2>" 				& trocar(Ucase( Visitante_Telefones(0)(5) )) & "</sms_2>")
					
					If Ubound(Visitante_Telefones) > 1 Then
						oFiletxt.WriteLine("<tipo_tel_3>" 			& trocar(Ucase( Visitante_Telefones(1)(0) )) & "</tipo_tel_3>")
						oFiletxt.WriteLine("<ddi_3>" 				& trocar(Ucase( Visitante_Telefones(1)(1) )) & "</ddi_3>")
						oFiletxt.WriteLine("<ddd_3>" 				& trocar(Ucase( Visitante_Telefones(1)(2) )) & "</ddd_3>")
						oFiletxt.WriteLine("<fone_3>" 				& limpar_formatacao(trocar(Ucase( Visitante_Telefones(1)(3) ))) & "</fone_3>")
						oFiletxt.WriteLine("<ramal_3>" 				& trocar(Ucase( Visitante_Telefones(1)(4) )) & "</ramal_3>")
						oFiletxt.WriteLine("<sms_3>" 				& trocar(Ucase( Visitante_Telefones(1)(5) )) & "</sms_3>")
					Else
						oFiletxt.WriteLine("<tipo_tel_3></tipo_tel_3>")
						oFiletxt.WriteLine("<ddi_3></ddi_3>")
						oFiletxt.WriteLine("<ddd_3></ddd_3>")
						oFiletxt.WriteLine("<fone_3></fone_3>")
						oFiletxt.WriteLine("<ramal_3></ramal_3>")
						oFiletxt.WriteLine("<sms_3></sms_3>")
					End If
										
					oFiletxt.WriteLine("<cargo>"				& trocar(Ucase( Cargo )) & "</cargo>")
					oFiletxt.WriteLine("<cargo_outros>" 		& trocar(Ucase( Cargo_Outros )) & "</cargo_outros>")
					oFiletxt.WriteLine("<subcargo>" 			& trocar(Ucase( SubCargo )) & "</subcargo>")
					oFiletxt.WriteLine("<subcargo_outros>" 		& trocar(Ucase( SubCargo_Outros )) & "</subcargo_outros>")
					oFiletxt.WriteLine("<departamento>" 		& trocar(Ucase( Depto )) & "</departamento>")
					oFiletxt.WriteLine("<departamento_outros>" 	& trocar(Ucase( Depto_Outros )) & "</departamento_outros>")
					
					oFiletxt.WriteLine("<cep>" 					& limpar_formatacao(trocar(Ucase( Endereco_CEP ))) & "</cep>")
					oFiletxt.WriteLine("<endereco>" 			& trocar(Ucase( Endereco_Endereco )) & "</endereco>")
					oFiletxt.WriteLine("<nro>" 					& trocar(Ucase( Endereco_Numero )) & "</nro>")
					oFiletxt.WriteLine("<complemento>" 			& trocar(Ucase( Endereco_Complemento )) & "</complemento>")
					oFiletxt.WriteLine("<bairro>" 				& trocar(Ucase( Endereco_Bairro )) & "</bairro>")
					oFiletxt.WriteLine("<cidade>" 				& trocar(Ucase( Endereco_Cidade )) & "</cidade>")
					oFiletxt.WriteLine("<uf>" 					& trocar(Ucase( Endereco_UF )) & "</uf>")
					oFiletxt.WriteLine("<pais>" 				& trocar(Ucase( Endereco_Pais )) & "</pais>")

					oFiletxt.WriteLine("<pesquisa>")
					If qtde_perguntas > 0 Then
						For w = Lbound(Perguntas_e_Respostas)+1 To Ubound(Perguntas_e_Respostas)
							oFiletxt.WriteLine("<pergunta questao='" & trocar(Perguntas_e_Respostas(w)(0)) & "' resposta='" & trocar(Perguntas_e_Respostas(w)(1)) & "'/>")
						Next
					End If
					oFiletxt.WriteLine("</pesquisa>")
				oFiletxt.WriteLine("</cadastro>")
		
			
				%>
                <%=Zeros_ESQ(qtos_zeros,x+1)%> - <b>IDC:</b> <%=id_cadastro%> / <b>IDV:</b> <%=id_visitante%> / <b>CPF:</b> <%=CPF%> / <b>Nome:</b> <%=nome_completo%> / <b>Emp.:</b> <%=Empresa_Nome_Fantasia%><br>
                <script language="javascript">document.getElementById('conteudo').scrollTop += 100;</script>
                <%
				response.Flush()
				
				SQL_Exportado = "Update Relacionamento_Cadastro " &_
								"Set " &_
								"	exportado = 1 " &_
								"	,Exportacao_DATA = getDate() " &_
								"	,Exportado_por_ID_Admin = " & Session("admin_id_usuario") & " " &_
								"Where ID_Relacionamento_Cadastro = " & id_cadastro
				Set RS_Exportado = Server.CreateObject("ADODB.RecordSet")
				RS_Exportado.Open SQL_Exportado, Conexao


				RS_2_Visitantes.Close
				Set RS_2_Visitantes = Nothing
			End If
			'======================================================
		Next
		'======================================================

		oFiletxt.WriteLine("</credenciamento>")
		oFiletxt.Close
		
		SQL_Arquivos = 	"Insert Into Arquivos_XML " &_
						"(arquivo, total, Id_Edicao) " &_
						"values " &_
						"('" & filename & "'," & total & "," & id_edicao & ")"
		Set RS_Arquivos = Server.CreateObject("ADODB.RecordSet")
		RS_Arquivos.Open SQL_Arquivos, Conexao
		%>      
        </div>
		<br><br>
        <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:left; width:15px; height:17px; float:left;"></div>
        <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco "><a href="evento.asp?id=<%=id_edicao%>" style="color: #fff">Listar Arquivos</a></div>
        <div style="background-image:url(/admin/images/bt_fundo.gif); background-position:right; width:15px; height:17px; float:left;"></div>
	<%
	EndTime = Timer()
	response.write("<br><br><span style='padding-left:50px;'>Tempo de processamento: " & FormatNumber((EndTime - StartTime),2) & " segundos</span>")
	
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
	data_termino = d & "-" & m & "-" & a & "_" & h & "-" & n & "-" & s
	%>
        <hr>
        Arquivo <B><%=Filename%></B> criado com sucesso<br><br>
        <big style="background-color:#FF0;"><a href="arquivos_<%=sAnoPasta%>/<%=Filename%>" target="_blank">Clique aqui * para salvar o arquivo.</a></big>&nbsp;* Bot�o direito > Salvar Como	<br>
        Total de Cadastros Listados : <b><%=total%></b><br>
        <br>
        Data Inicio: <%=data_inicio%><br>
        Data Termino: <%=data_termino%>
        <script>window.scrollMaxY()</script>
    </td>
  </tr>
</table>
<% response.Flush() %>
<table width="955" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="/admin/images/img_tabela_branca_inferior.jpg" width="955" height="15" /></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>

</body>
</html>
<% Conexao.Close %>