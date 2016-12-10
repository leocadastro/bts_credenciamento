<%
sub rw(s)
	response.Write("<textarea style='width:300px; height:250px;'>"&s&"</textarea>")
end sub

DIM StartTime
Dim EndTime
StartTime = Timer()
%>

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

'==================================================================
		id_edicao = Limpar_Texto(request("id_edicao"))
		id_idioma = Limpar_Texto(request("id_idioma"))
		id_tipo = Limpar_Texto(request("id_tipo"))
'==================================================================

Function RemoverAcentuacao (texto)
	limpar = Ucase(texto)
	If Len(limpar) <= 0 Then
	Else
		limpar = Replace(limpar, "Á", "A")
		limpar = Replace(limpar, "À", "A")
		limpar = Replace(limpar, "Â", "A")
		limpar = Replace(limpar, "Á", "A")
		limpar = Replace(limpar, "Ã", "A")
		limpar = Replace(limpar, "É", "E")
		limpar = Replace(limpar, "È", "E")
		limpar = Replace(limpar, "Ê", "E")
		limpar = Replace(limpar, "Í", "I")
		limpar = Replace(limpar, "Ì", "I")
		limpar = Replace(limpar, "Î", "I")
		limpar = Replace(limpar, "Ó", "O")
		limpar = Replace(limpar, "Ò", "O")
		limpar = Replace(limpar, "Ô", "O")
		limpar = Replace(limpar, "Õ", "O")
		limpar = Replace(limpar, "Ú", "U")
		limpar = Replace(limpar, "Ù", "U")
		limpar = Replace(limpar, "Û", "U")
		limpar = Replace(limpar, "Ç", "Ç")
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
							"From  Eventos_Edicoes as EE " &_
							"Inner Join Eventos as E ON E.ID_Evento = EE.ID_Evento " &_
							"Where  " &_
							"	Ee.ID_Edicao = " & id_edicao
	'rw SQL_Evento_Autorizado
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
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">Evento: <%=Feira%></span></td>
        <td width="100" align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;">&nbsp;</td>
      </tr>
    </table>
      <div id="aviso" style="background-color:#FFFF00; width:100%; text-align:center;" class="t_arial fs11px bold c_preto"> <img src="/admin/images/ico_aviso.gif" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso"> <span id="aviso_conteudo">Aviso Processando, por favor não interrompa.</span></div>

		<%
		response.Flush()
		
		Where_Cadastros = ""
		Tipo_Arquivo = ""
		If Len(Trim(id_idioma)) > 0 Then 
			Where_Cadastros = Where_Cadastros & " AND RC.ID_Idioma = " & id_idioma
			Tipo_Arquivo	= Tipo_Arquivo & "_Idioma(" & id_idioma & ")_"
		End IF
		If Len(Trim(id_tipo)) > 0 Then 
			Where_Cadastros = Where_Cadastros & " AND TC.ID_Formulario = " & id_tipo
			Tipo_Arquivo	= Tipo_Arquivo & "_Formulario(" & id_tipo & ")_"
			Where_Perguntas = "AND P.ID_Formulario = " & id_tipo & " " 
		End If

		SQL_1_RelacionamentoCadastros =		"SELECT " &_
											"	RC.ID_Relacionamento_Cadastro " &_
											"	,RC.ID_Idioma " &_
											"	,RC.ID_Tipo_Credenciamento " &_
											"	,RC.ID_Visitante " &_
											"	,RC.ID_Empresa " &_
											"	,RC.ID_Edicao " &_
											"	,TC.ID_Formulario " &_
											"	,RC.Data_Cadastro " &_
											"FROM Relacionamento_Cadastro AS RC  " &_
											"INNER JOIN Tipo_Credenciamento AS TC  " &_
											"	ON TC.ID_Tipo_Credenciamento = RC.ID_Tipo_Credenciamento " &_
											"WHERE     (RC.Ativo = 1) AND (RC.ID_Visitante IS NOT NULL) AND (RC.ID_Edicao = " & id_edicao & ") " &_
											" " & Where_Cadastros
											
		'rw SQL_1_RelacionamentoCadastros

		'response.write("<hr><b>SQL_1_RelacionamentoCadastros</b><br>" & SQL_1_RelacionamentoCadastros & "<br>")

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
		<br><span style=''>Tempo da 1ª Listagem: <%=FormatNumber((Intermediaria1 - StartTime),2)%> segundos</span><br>
        <b>Média</b> de tempo para cada 100 registros: 13 segundos.<br>
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
						
		extensao = ".XLS"
	
		Filename = arquivo & "_" & RemoverAcentuacao(feira) & Tipo_Arquivo & "_" & data & extensao ' file to read
	
		Const ForReading = 1, ForWriting = 2, ForAppending = 3
		Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
	
		' Create a filesystem object
		Dim FSO
		set FSO = server.createObject("Scripting.FileSystemObject")
	
		' Map the logical path to the physical system path
		Dim Filepath
		Filepath = Server.MapPath("/admin/exportar_xml/arquivos_2012/" & Filename)
	
		Set oFiletxt = FSO.CreateTextFile(Filepath, True)
		sPath = FSO.GetAbsolutePathName(Filepath)
		sFilename = FSO.GetFileName(sPath)
	
		oFiletxt.WriteLine("<table border=1>")
		oFiletxt.WriteLine("<tr>")
		oFiletxt.WriteLine("<td rowspan='3'>N</td>")
		oFiletxt.WriteLine("<td rowspan='3'>ID_Cadastro</td>")
		oFiletxt.WriteLine("<td rowspan='3'>ID_Visitante</td>")
		oFiletxt.WriteLine("<td rowspan='3'>ID_Empresa</td>")
		oFiletxt.WriteLine("<td rowspan='3'>ID_Formulario</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Data_Cadastro</td>")
		
		oFiletxt.WriteLine("<td rowspan='3'>Empresa_CNPJ</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Empresa_Razao_Social</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Empresa_Nome_Fantasia</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Empresa_Qtde_Funcionarios</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Empresa_Principal_Produto</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Ramos e Atividades</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Ramos Outros</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Atividades Outros</td>")
		oFiletxt.WriteLine("<td rowspan='3'>InteresseFeira</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Empresa_Site</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Empresa_Presidente</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Empresa_Reitor</td>")

		oFiletxt.WriteLine("<td rowspan='3'>Empresa_Tipo_Telefone</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Empresa_DDI</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Empresa_DDD</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Empresa_Numero</td>")					
		oFiletxt.WriteLine("<td rowspan='3'>Empresa_Ramal</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Empresa_SMS</td>")

		oFiletxt.WriteLine("<td rowspan='3'>CPF</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Passaporte</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Nome_Completo</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Nome_Credencial</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Data_Nasc</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Sexo</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Email</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Newsletter</td>")
		
		oFiletxt.WriteLine("<td rowspan='3'>Visitante_Tipo_Telefone 1</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Visitante_DDI 1</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Visitante_DDD 1</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Visitante_Numero 1</td>")					
		oFiletxt.WriteLine("<td rowspan='3'>Visitante_Ramal 1</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Visitante_SMS 1</td>")
		
		oFiletxt.WriteLine("<td rowspan='3'>Visitante_Tipo_Telefone 2</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Visitante_DDI 2</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Visitante_DDD 2</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Visitante_Numero 2</td>")					
		oFiletxt.WriteLine("<td rowspan='3'>Visitante_Ramal 2</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Visitante_SMS 2</td>")
									
		oFiletxt.WriteLine("<td rowspan='3'>Cargo</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Cargo_Outros</td>")
		oFiletxt.WriteLine("<td rowspan='3'>SubCargo</td>")
		oFiletxt.WriteLine("<td rowspan='3'>SubCargo_Outros</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Depto</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Depto_Outros</td>")
		
		oFiletxt.WriteLine("<td rowspan='3'>Endereco_CEP</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Endereco_Endereco</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Endereco_Numero</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Endereco_Complemento</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Endereco_Bairro</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Endereco_Cidade</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Endereco_UF</td>")
		oFiletxt.WriteLine("<td rowspan='3'>Endereco_Pais</td>")
		

		SQL_1_Perguntas_Itens =		"Select " &_
											"	REP.ID_Relacionamento_Pergunta, " &_
											"	REP.ID_Pergunta, " &_
											"	P.ID_Formulario, " &_
											"	P.Pergunta_PTB, " &_
											"	P.Tipo, " &_
											"	P.Multiplo " &_
											"From  " &_
											"	Relacionamento_Edicoes_Pergunta as REP " &_
											"Inner Join " &_
											"	Perguntas as P " &_
											"	On REP.ID_Pergunta = P.ID_Perguntas " &_
											"Where  " &_
											"	REP.ID_Edicao = " & id_edicao & " " &_
											"	And P.Ativo = 1 " &_
											" " & Where_Perguntas 

		response.write("<hr><b>SQL_1_Perguntas_Itens</b><br>" & SQL_1_Perguntas_Itens & "<br>")

		Set RS_1_Perguntas_Itens = Server.CreateObject("ADODB.RecordSet")
		RS_1_Perguntas_Itens.CursorLocation = 3
		RS_1_Perguntas_Itens.CursorType = 3
		RS_1_Perguntas_Itens.LockType = 1
		RS_1_Perguntas_Itens.Open SQL_1_Perguntas_Itens, Conexao

		If not RS_1_Perguntas_Itens.BOF or not RS_1_Perguntas_Itens.EOF Then

			colspan_cabecalho1 = 0
			While not RS_1_Perguntas_Itens.EOF
				colspan_cabecalho1 = colspan_cabecalho1 + 1
				RS_1_Perguntas_Itens.MoveNext
			Wend
			RS_1_Perguntas_Itens.MoveFirst

			oFiletxt.WriteLine("<td colspan='" & colspan_cabecalho1 & "'>Perguntas e Respostas</td>")

			SQL_TipoFormularios = 	"Select " &_
									"	Count(P.ID_Formulario) as ColSpan_Tipo " &_
									"	,P.ID_Formulario " &_
									"	,F.Nome as Formulario " &_
									"From  " &_
									"	Relacionamento_Edicoes_Pergunta as REP " &_
									"Inner Join " &_
									"	Perguntas as P " &_
									"	On REP.ID_Pergunta = P.ID_Perguntas " &_
									"Inner Join Formularios as F " &_
									"	On F.ID_Formulario = REP.ID_Formulario " &_
									"Where  " &_
									"	REP.ID_Edicao = " & id_edicao & " " &_
									"	And P.Ativo = 1 " &_
									"Group by P.ID_Formulario, F.Nome"

			Set RS_TipoFormularios = Server.CreateObject("ADODB.RecordSet")
			RS_TipoFormularios.CursorLocation = 3
			RS_TipoFormularios.CursorType = 3
			RS_TipoFormularios.LockType = 1
			RS_TipoFormularios.Open SQL_TipoFormularios, Conexao

			Tipo_Formulario_Perguntas = ""
			Qtde_Formulario_Perguntas = ""
			Nome_Formulario_Perguntas = ""

			' Escreve os tipos de Formularios'
			If not RS_TipoFormularios.BOF or not RS_TipoFormularios.EOF Then
				oFiletxt.WriteLine("<tr>")
				While not RS_TipoFormularios.EOF
					' Escrever o Cabecalho
					Tipo_Formulario_Perguntas = RS_TipoFormularios.Fields("ID_Formulario")
					Qtde_Formulario_Perguntas = RS_TipoFormularios.Fields("ColSpan_Tipo")
					Nome_Formulario_Perguntas = RS_TipoFormularios.Fields("Formulario")
					oFiletxt.WriteLine("<td colspan='" & Qtde_Formulario_Perguntas & "' nowrap>" & Nome_Formulario_Perguntas & "</td>")
					RS_TipoFormularios.MoveNext
				Wend
				oFiletxt.WriteLine("</tr>")
				RS_TipoFormularios.Close
				Set RS_TipoFormularios = Nothing
			End If

			' Escreve as Perguntas '
			If not RS_1_Perguntas_Itens.BOF or not RS_1_Perguntas_Itens.EOF Then
				oFiletxt.WriteLine("<tr>")
				Redim Array_ID_Perguntas(colspan_cabecalho1)
				i_id_perguntas = 0
				While not RS_1_Perguntas_Itens.EOF
					i_id_perguntas = i_id_perguntas + 1
					oFiletxt.WriteLine("<td nowrap>" & RS_1_Perguntas_Itens.Fields("Pergunta_PTB").Value & "</td>")
					Array_ID_Perguntas(i_id_perguntas) = RS_1_Perguntas_Itens.Fields("ID_Pergunta").Value
					RS_1_Perguntas_Itens.MoveNext
				Wend
				oFiletxt.WriteLine("</tr>")
			End If
			RS_1_Perguntas_Itens.Close
			Set RS_1_Perguntas_Itens = Nothing
		End If
			
		oFiletxt.WriteLine("</tr>")
		total = 0
		'======================================================
		For x = Lbound(IDs) To Ubound(IDs) - 1
		
			'======================================================
			SQL_2_Visitantes = 	"Select " &_
								"	* " &_
								"From Vw_Visitantes " &_
								"Where  " &_
								"	ID_Visitante = " & IDs(x)(3)
'response.write("<hr><b>SQL_2_Visitantes</b><br>" & SQL_2_Visitantes & "<br>")\
				
			Set RS_2_Visitantes = Server.CreateObject("ADODB.RecordSet")
			RS_2_Visitantes.CursorLocation = 3
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
				RS_3_Telefones.CursorLocation = 3
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
				RS_6_Empresa.CursorLocation = 3
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
				RS_7_Empresa_Produtos.CursorLocation = 3
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
										
				'Response.Write("<hr>SQL_8_Empresa_Ramos:" & SQL_8_Empresa_Ramos & "<hr>")
					
				Set RS_8_Empresa_Ramos = Server.CreateObject("ADODB.RecordSet")
				RS_8_Empresa_Ramos.CursorLocation = 3
				RS_8_Empresa_Ramos.CursorType = 3
				RS_8_Empresa_Ramos.LockType = 1
				RS_8_Empresa_Ramos.Open SQL_8_Empresa_Ramos, Conexao				
				'======================================================
				
				'======================================================
				Ramos_e_Atividades = ""
				Ramos_Outros = ""
				Atividadades_Outros = ""
				
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
				'If not RS_8_Empresa_Ramos.BOF or not RS_8_Empresa_Ramos.EOF Then
'					While not RS_8_Empresa_Ramos.EOF
'						Ramo			= RS_8_Empresa_Ramos.Fields("Ramo").Value
'						Ramo_Outro		= RS_8_Empresa_Ramos.Fields("Ramo_Outros").Value
'						Atividade		= RS_8_Empresa_Ramos.Fields("Atividade").Value
'						Atividade_Outro	= RS_8_Empresa_Ramos.Fields("Atividade_Outros").Value
'						' Se virer Outros
'						If Len(Trim(Ramo_Outro)) > 0  Then
'							If Len(Trim(Ramos_Outros)) > 0 Then
'								Ramos_Outros = Ramos_Outros & "; "
'							End If
'							Ramos_Outros 		= Ramos_Outros & " " & Ramos_Outros
'						Else
'							' Nao exibir OUTROS
''							If Ramo <> "Outros" Then
'								Ramos_e_Atividades 		= Ramos_e_Atividades & Ramo & " - " & Atividade
''							End If
'						End If
'						If Len(Trim(Atividade_Outro)) > 0  Then
'							If Len(Trim(Atividadades_Outros)) > 0 Then
'								Atividadades_Outros = Atividadades_Outros & "; "
'							End If
'							Atividadades_Outros		= Atividadades_Outros & " " & Atividade_Outro
'						End If
'						RS_8_Empresa_Ramos.MoveNext
'						If not RS_8_Empresa_Ramos.EOF Then Ramos_e_Atividades = Ramos_e_Atividades & "; "
'					Wend
'					RS_8_Empresa_Ramos.Close
'					Set RS_8_Empresa_Ramos = Nothing
'				End If
				'======================================================
				
				
				'======================================================
				SQL_10_Empresa_InteresseFeira = 	"Select " &_
													"	* " &_
													"From Vw_InteresseFeira " &_
													"Where  " &_
													"	ID_Relacionamento_Cadastro = " & IDs(x)(0)
					
				Set RS_10_Empresa_InteresseFeira = Server.CreateObject("ADODB.RecordSet")
				RS_10_Empresa_InteresseFeira.CursorLocation = 3
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
			RS_5_Endereco.CursorLocation = 3
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
				idioma 			= IDs(x)(1)
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
				RS_4_Telefones.CursorLocation = 3
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

				' MASCARA DE CNPJ
				If Len(Trim(Empresa_CNPJ)) > 0 Then
					Empresa_CNPJ 		= limpar_formatacao(trocar(Ucase( Empresa_CNPJ )))
					Empresa_CNPJ_MASK 	= Mid(Empresa_CNPJ,1,2) & "." & Mid(Empresa_CNPJ,3,3) & "." & Mid(Empresa_CNPJ,6,3) & "/" & Mid(Empresa_CNPJ,9,4) & "-" & Mid(Empresa_CNPJ,13,2)
				Else
					Empresa_CNPJ_MASK = ""
				End If

				' MASCARA DE CPF
				IF Len(Trim(CPF)) > 0 Then
					CPF 		 		= limpar_formatacao(trocar(Ucase( CPF )))
					CPF_MASK 			= Mid(CPF,1,3) & "." & Mid(CPF,4,3) & "." & Mid(CPF,7,3) & "-" & Mid(CPF,10,2)
				Else
					CPF_MASK = ""
				End If 

				' MASCARA DE DATA NASCIMENTO
				Data_Nasc 			= limpar_formatacao(trocar(Ucase( Data_Nasc )))
				Data_Nasc_MASK		= Mid(Data_Nasc,1,2) & "/" & Mid(Data_Nasc,3,2) & "/" & Mid(Data_Nasc,5,4)

				' MASCARA DE CEP
				If Len(Trim(Endereco_CEP)) > 0 And idioma = "1" Then
					Endereco_CEP		= limpar_formatacao(trocar(Ucase( Endereco_CEP )))
					Endereco_CEP_MASK	= Mid(Endereco_CEP,1,5) & "-" & Mid(Endereco_CEP,6,3)
				Elseif Len(Trim(Endereco_CEP)) > 0 And idioma <> "1" Then
					Endereco_CEP_MASK 	= "'" & limpar_formatacao(trocar(Ucase( Endereco_CEP )))
				Else
					Endereco_CEP_MASK = "" 
				End If
				
				oFiletxt.WriteLine("<tr>")
					oFiletxt.WriteLine("<td nowrap>" & i & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( ID_Cadastro )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( ID_Visitante )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( id_empresa )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( formulario )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( data_cadastro )) & "</td>")
					
					oFiletxt.WriteLine("<td nowrap>" & Empresa_CNPJ_MASK & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Empresa_Razao_Social )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Empresa_Nome_Fantasia )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Empresa_Qtde_Funcionarios )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Empresa_Principal_Produto )) & "</td>")
					
					
					If qtde_ramos > 0 Then
						w=0
						'response.Write("qtde_ramos="&qtde_ramos&"<br>")
						while w < Ubound(Ramos_Atividades)
							arrRamAtiv = split(RemoverAcentuacao(Ramos_Atividades(w)(0))," - ")
							oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( arrRamAtiv(0) )) & "</td>")
							oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( arrRamAtiv(1) )) & "</td>")
							w=w+1
						wend
					End If
					

					
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Atividadades_Outros )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( InteresseFeira )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Empresa_Site )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Empresa_Presidente )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Empresa_Reitor )) & "</td>")

					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Empresa_Tipo_Telefone )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Empresa_DDI )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Empresa_DDD )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & limpar_formatacao(trocar(Ucase( Empresa_Numero ))) & "</td>")					
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Empresa_Ramal )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Empresa_SMS )) & "</td>")

					oFiletxt.WriteLine("<td nowrap>" & CPF_MASK & "</td>")
					oFiletxt.WriteLine("<td nowrap>'" & trocar(Ucase( Passaporte )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Nome_Completo )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Nome_Credencial )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & Data_Nasc_MASK & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Sexo )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Email )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Newsletter )) & "</td>")

					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Visitante_Telefones(0)(0) )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Visitante_Telefones(0)(1) )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Visitante_Telefones(0)(2) )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & limpar_formatacao(trocar(Ucase( Visitante_Telefones(0)(3) ))) & "</td>")					
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Visitante_Telefones(0)(4) )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Visitante_Telefones(0)(5) )) & "</td>")
					
					If Ubound(Visitante_Telefones) > 1 Then
						oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Visitante_Telefones(1)(0) )) & "</td>")
						oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Visitante_Telefones(1)(1) )) & "</td>")
						oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Visitante_Telefones(1)(2) )) & "</td>")
						oFiletxt.WriteLine("<td nowrap>" & limpar_formatacao(trocar(Ucase( Visitante_Telefones(1)(3) ))) & "</td>")
						oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Visitante_Telefones(1)(4) )) & "</td>")
						oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Visitante_Telefones(1)(5) )) & "</td>")
					Else
						oFiletxt.WriteLine("<td>&nbsp;</td>")
						oFiletxt.WriteLine("<td>&nbsp;</td>")
						oFiletxt.WriteLine("<td>&nbsp;</td>")
						oFiletxt.WriteLine("<td>&nbsp;</td>")
						oFiletxt.WriteLine("<td>&nbsp;</td>")
						oFiletxt.WriteLine("<td>&nbsp;</td>")
					End If
										
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Cargo )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Cargo_Outros )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( SubCargo )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( SubCargo_Outros )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Depto )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Depto_Outros )) & "</td>")
					
					oFiletxt.WriteLine("<td nowrap>" & Endereco_CEP_MASK & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Endereco_Endereco )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Endereco_Numero )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Endereco_Complemento )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Endereco_Bairro )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Endereco_Cidade )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Endereco_UF )) & "</td>")
					oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( Endereco_Pais )) & "</td>")
	
					' Loop das Repostas '
					For id_perg = Lbound(Array_ID_Perguntas)+1 To Ubound(Array_ID_Perguntas)

						SQL_Listar_Resposta_Pergunta_por_Item = 	"Select Distinct " &_
																	"	PO.Opcao_PTB as Resposta, " &_
																	"	RP.Texto as Complemento " &_
																	"From Relacionamento_Perguntas as RP " &_
																	"Inner Join Perguntas_Opcoes as PO " &_
																	"	On PO.ID_Opcoes = RP.ID_Opcoes " &_
																	"Where  " &_
																	"	RP.ID_Relacionamento_Cadastro = " & IDs(x)(0) & " " &_
																	"	AND RP.ID_Perguntas = " & Array_ID_Perguntas(id_perg)

						'response.write("<b>Ubound(Array_ID_Perguntas)</b><br>" & Ubound(Array_ID_Perguntas))
						'response.write("<hr><b>SQL_Listar_Resposta_Pergunta_por_Item</b><br>" & SQL_Listar_Resposta_Pergunta_por_Item)

						Set RS_Listar_Resposta_Pergunta_por_Item = Server.CreateObject("ADODB.RecordSet")
						RS_Listar_Resposta_Pergunta_por_Item.CursorLocation = 3
						RS_Listar_Resposta_Pergunta_por_Item.CursorType = 3
						RS_Listar_Resposta_Pergunta_por_Item.LockType = 1
						RS_Listar_Resposta_Pergunta_por_Item.Open SQL_Listar_Resposta_Pergunta_por_Item, Conexao

						resposta_a_esta_pergunta = ""
						If not RS_Listar_Resposta_Pergunta_por_Item.BOF or not RS_Listar_Resposta_Pergunta_por_Item.EOF Then
							While not RS_Listar_Resposta_Pergunta_por_Item.EOF
								resposta_a_esta_pergunta = resposta_a_esta_pergunta & RS_Listar_Resposta_Pergunta_por_Item.Fields("Resposta").Value
								' Se houver complemento, coloque : + complemento'
								If Len(Trim(RS_Listar_Resposta_Pergunta_por_Item.Fields("Complemento").Value)) > 0 Then
									resposta_a_esta_pergunta = resposta_a_esta_pergunta & ": " & RS_Listar_Resposta_Pergunta_por_Item.Fields("Complemento").Value
								End If
								RS_Listar_Resposta_Pergunta_por_Item.MoveNext

								' Se não ACABOU colocar ;'
								If not RS_Listar_Resposta_Pergunta_por_Item.EOF Then
									resposta_a_esta_pergunta = resposta_a_esta_pergunta & "; "
								End If
							Wend
							RS_Listar_Resposta_Pergunta_por_Item.Close
							Set RS_Listar_Resposta_Pergunta_por_Item = Nothing
						End If
						oFiletxt.WriteLine("<td nowrap>" & trocar(Ucase( resposta_a_esta_pergunta )) & "</td>")
					Next
				oFiletxt.WriteLine("</tr>")
		
			
				%>
                <%=Zeros_ESQ(qtos_zeros,x+1)%> - <b>IDIOMA:</b> <%=idioma%> / <b>IDC:</b> <%=id_cadastro%> / <b>IDV:</b> <%=id_visitante%> / <b>CPF:</b> <%=CPF%> / <b>Nome:</b> <%=nome_completo%><br>
                <script language="javascript">document.getElementById('conteudo').scrollTop += 100;</script>
                <%
				response.Flush()

				RS_2_Visitantes.Close
				Set RS_2_Visitantes = Nothing
			End If
			'======================================================
		Next
		'======================================================

		oFiletxt.WriteLine("</table>")
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
        <div style="background-image:url(/admin/images/bg_bt_fundo.gif); background-position:center; height:17px; float:left; text-align:center; line-height:17px;" class="fs10px t_verdana c_branco "><a href="default.asp?id=<%=id_edicao%>" style="color: #fff">Listar Arquivos</a></div>
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
        <big style="background-color:#FF0;"><a href="/admin/exportar_xml/arquivos_2012/<%=Filename%>" target="_blank">Clique aqui * para salvar o arquivo.</a></big>&nbsp;* Botão direito > Salvar Como	<br>
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