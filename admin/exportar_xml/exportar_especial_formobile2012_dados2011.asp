<%
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

id_edicao	= "2"

Function RemoverAcentuacao (texto)
	limpar = Ucase(texto)
	If Len(limpar) <= 0 or Isnull(limpar) Then
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
		limpar = Replace(limpar, "Ç", "C")
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
		limpar = Replace(limpar, "(", "")
		limpar = Replace(limpar, ")", "")
		If IsNumeric(limpar) = False Then limpar = ""
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
Conexao.Open	"DRIVER=SQL Server;UID=credenciamento;PWD=cred#btsp;DATABASE=CredenciamentoBTS;SERVER=SABR-DATABASE;Persist Security Info=True;"
'==================================================

'==================================================
Set Conexao_2012 = Server.CreateObject("ADODB.Connection")
Conexao_2012.Open	Application("cnn")
'==================================================

	SQL_Evento_Autorizado = "Select " &_
							"	EE.Ano " &_
							"	,E.Nome_PTB as Feira " &_
							"From Administradores_Edicoes AE " &_
							"Inner Join Eventos_Edicoes as EE ON Ae.ID_Edicao = EE.ID_Edicao " &_
							"Inner Join Eventos as E ON E.ID_Evento = EE.ID_Evento " &_
							"Where  " &_
							"	Ae.ID_Admin = 3 " &_
							"	AND Ae.ID_Edicao = " & id_edicao
	Set RS_Evento_Autorizado = Server.CreateObject("ADODB.Recordset")
	RS_Evento_Autorizado.Open SQL_Evento_Autorizado, Conexao_2012
	
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
        <td valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;" align="center"><span style="color: #B01D22">Exportar XML<br>
          Evento: <%=Feira%></span></td>
        <td width="100" align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:18px; font-weight:bold;"><a href="evento.asp?id=<%=id_edicao%>"><img src="/admin/images/bt_voltar_admin.gif" width="100" height="48" border="0" /></a></td>
      </tr>
    </table>
      <div id="aviso" style="background-color:#FFFF00; width:100%; text-align:center;" class="t_arial fs11px bold c_preto"> <img src="/admin/images/ico_aviso.gif" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso"> <span id="aviso_conteudo">Aviso Processando, por favor não interrompa.</span></div>

		<%
		response.Flush()

		SQL_1_RelacionamentoCadastros =		"Select  " &_
											"	id " &_
											"From ForMobile_2010 " &_
											"Where  " &_
											"	Exportado is null "

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
				id 			= RS_1_RelacionamentoCadastros.Fields("id").Value
				IDs(i) = id
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
	
		arquivo = "pre_credenciados_especial_dados2011"
						
		extensao = ".xml"
	
		Filename = arquivo & "_" & RemoverAcentuacao(feira) & "_" & data & extensao ' file to read
	
		Const ForReading = 1, ForWriting = 2, ForAppending = 3
		Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
	
		' Create a filesystem object
		Dim FSO
		set FSO = server.createObject("Scripting.FileSystemObject")
	
		' Map the logical path to the physical system path
		Dim Filepath
		Filepath = Server.MapPath("arquivos_2012/" & Filename)
	
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
								"	id,  " &_
								"	idioma,  " &_
								"	txt_cpf,  " &_
								"	txt_nome,  " &_
								"	txt_endereco,  " &_
								"	txt_numero,  " &_
								"	txt_complemento,  " &_
								"	txt_bairro,  " &_
								"	txt_cidade,  " &_
								"	txt_estado,  " &_
								"	txt_cep,  " &_
								"	txt_telefone,  " &_
								"	txt_fax,  " &_
								"	txt_celular,  " &_
								"	txt_email,  " &_
								"	txt_site,  " &_
								"	txt_cnpj,  " &_
								"	txt_empresa,  " &_
								"	txt_cargo,  " &_
								"	txt_decisao_compra,  " &_
								"	txt_ramo_atividade,  " &_
								"	txt_qtde_funcionarios,  " &_
								"	txt_ficou_sabendo_via,  " &_
								"	bit_email_marketing,  " &_
								"	exportado,  " &_
								"	dt_exportado,  " &_
								"	dt_cadastro " &_
								"From ForMobile_2010 " &_ 
								"Where id = " & IDs(x)
'response.write("<hr><b>SQL_2_Visitantes</b><br>" & SQL_2_Visitantes & "<br>")
				
			Set RS_ForMobile = Server.CreateObject("ADODB.RecordSet")
			RS_ForMobile.CursorLocation = 2
			RS_ForMobile.CursorType = 3
			RS_ForMobile.LockType = 1
			RS_ForMobile.Open SQL_2_Visitantes, Conexao
			'======================================================
			
'			response.write("Data: " & RS_ForMobile("dt_cadastro") & "<br>")
			
			id_cadastro = IDs(x)
				'======================================================
				total = total + 1
							
				oFiletxt.WriteLine("<cadastro>")
					oFiletxt.WriteLine("<id_cadastro>" 			& trocar(Ucase( ID_Cadastro )) & "</id_cadastro>")
					oFiletxt.WriteLine("<id_visitante></id_visitante>")
					oFiletxt.WriteLine("<id_empresa></id_empresa>")
					oFiletxt.WriteLine("<formulario></formulario>")
					oFiletxt.WriteLine("<data_cadastro>" 		& trocar(Ucase( RS_ForMobile("dt_cadastro") )) & "</data_cadastro>")
					
					oFiletxt.WriteLine("<cnpj>" 				& limpar_formatacao(trocar(Ucase( RS_ForMobile("txt_cnpj") ))) & "</cnpj>")
					oFiletxt.WriteLine("<razao_social>"			& trocar(Ucase( RS_ForMobile("txt_empresa") )) & "</razao_social>")
					oFiletxt.WriteLine("<nome_fantasia>"		& trocar(Ucase( RS_ForMobile("txt_empresa") )) & "</nome_fantasia>")
					oFiletxt.WriteLine("<qtde_funcionarios>"	& trocar(Ucase( RS_ForMobile("txt_qtde_funcionarios") )) & "</qtde_funcionarios>")
					oFiletxt.WriteLine("<principal_produto></principal_produto>")

'	MODELO ANTIGO					
'					oFiletxt.WriteLine("<ramos>"				& trocar(Ucase( Ramos )) & "</ramos>")
'					oFiletxt.WriteLine("<atividades>"			& trocar(Ucase( Atividadades )) & "</atividades>")

'	MODELO NOVO 12 07 2012 20:30 HOMERO
					oFiletxt.WriteLine("<ramos_atividades>")
							oFiletxt.WriteLine("<item ramo_atividade='" & RemoverAcentuacao(RS_ForMobile("txt_ramo_atividade")) & "' ramo_outros='' atividade_outros='' />")
					oFiletxt.WriteLine("</ramos_atividades>")
					
					oFiletxt.WriteLine("<interesses_feira></interesses_feira>")
					oFiletxt.WriteLine("<site>"					& trocar(Ucase( RS_ForMobile("txt_site") )) & "</site>")
					oFiletxt.WriteLine("<presidente></presidente>")
					oFiletxt.WriteLine("<reitor></reitor>")

					oFiletxt.WriteLine("<tipo_tel_1></tipo_tel_1>")
					oFiletxt.WriteLine("<ddi_1></ddi_1>")
					oFiletxt.WriteLine("<ddd_1></ddd_1>")
					oFiletxt.WriteLine("<fone_1>" 				& limpar_formatacao(trocar(Ucase( RS_ForMobile("txt_telefone") ))) & "</fone_1>")					
					oFiletxt.WriteLine("<ramal_1></ramal_1>")
					oFiletxt.WriteLine("<sms_1></sms_1>")

					oFiletxt.WriteLine("<cpf>" 					& limpar_formatacao(trocar(Ucase( RS_ForMobile("txt_cpf") ))) & "</cpf>")
					oFiletxt.WriteLine("<nome>" 				& trocar(Ucase( RS_ForMobile("txt_nome") )) & "</nome>")
					oFiletxt.WriteLine("<credencial>" 			& trocar(Ucase( RS_ForMobile("txt_nome") )) & "</credencial>")
					oFiletxt.WriteLine("<data_nasc></data_nasc>")
					oFiletxt.WriteLine("<sexo></sexo>")
					oFiletxt.WriteLine("<email>" 				& trocar(Ucase( RS_ForMobile("txt_email") )) & "</email>")
					oFiletxt.WriteLine("<newsletter>" 			& trocar(Ucase( RS_ForMobile("bit_email_marketing") )) & "</newsletter>")

					oFiletxt.WriteLine("<tipo_tel_2></tipo_tel_2>")
					oFiletxt.WriteLine("<ddi_2></ddi_2>")
					oFiletxt.WriteLine("<ddd_2></ddd_2>")
					oFiletxt.WriteLine("<fone_2>" 				& limpar_formatacao(trocar(Ucase( RS_ForMobile("txt_fax") ))) & "</fone_2>")					
					oFiletxt.WriteLine("<ramal_2></ramal_2>")
					oFiletxt.WriteLine("<sms_2></sms_2>")
					
					oFiletxt.WriteLine("<tipo_tel_3></tipo_tel_3>")
					oFiletxt.WriteLine("<ddi_3></ddi_3>")
					oFiletxt.WriteLine("<ddd_3></ddd_3>")
					oFiletxt.WriteLine("<fone_3>" 				& limpar_formatacao(trocar(Ucase( RS_ForMobile("txt_celular") ))) & "</fone_3>")
					oFiletxt.WriteLine("<ramal_3></ramal_3>")
					oFiletxt.WriteLine("<sms_3></sms_3>")
										
					oFiletxt.WriteLine("<cargo>"				& trocar(Ucase( RS_ForMobile("txt_cargo") )) & "</cargo>")
					oFiletxt.WriteLine("<cargo_outros></cargo_outros>")
					oFiletxt.WriteLine("<subcargo></subcargo>")
					oFiletxt.WriteLine("<subcargo_outros></subcargo_outros>")
					oFiletxt.WriteLine("<departamento></departamento>")
					oFiletxt.WriteLine("<departamento_outros></departamento_outros>")
					
					oFiletxt.WriteLine("<cep>" 					& limpar_formatacao(trocar(Ucase( RS_ForMobile("txt_cep") ))) & "</cep>")
					oFiletxt.WriteLine("<endereco>" 			& trocar(Ucase( RS_ForMobile("txt_endereco") )) & "</endereco>")
					oFiletxt.WriteLine("<nro>" 					& trocar(Ucase( RS_ForMobile("txt_numero") )) & "</nro>")
					oFiletxt.WriteLine("<complemento>" 			& trocar(Ucase( RS_ForMobile("txt_complemento") )) & "</complemento>")
					oFiletxt.WriteLine("<bairro>" 				& trocar(Ucase( RS_ForMobile("txt_bairro") )) & "</bairro>")
					oFiletxt.WriteLine("<cidade>" 				& trocar(Ucase( RS_ForMobile("txt_cidade") )) & "</cidade>")
					oFiletxt.WriteLine("<uf>" 					& trocar(Ucase( RS_ForMobile("txt_estado") )) & "</uf>")
					oFiletxt.WriteLine("<pais></pais>")

					oFiletxt.WriteLine("<pesquisa>")
					oFiletxt.WriteLine("</pesquisa>")
				oFiletxt.WriteLine("</cadastro>")
		
			
				%>
                <%=Zeros_ESQ(qtos_zeros,x+1)%> - <b>IDC:</b> <%=id_cadastro%> / <b>CPF:</b> <%=RS_ForMobile("txt_cpf")%> / <b>Nome:</b> <%=RS_ForMobile("txt_nome")%> / <b>Emp.:</b> <%=RS_ForMobile("txt_empresa")%><br>
                <script language="javascript">document.getElementById('conteudo').scrollTop += 100;</script>
                <%
				response.Flush()
				
				SQL_Exportado = "Update ForMobile_2010 " &_
								"	Set exportado = 1, " &_
								"	dt_exportado = getDate() " &_
								"Where id = " & id_cadastro 
				Set RS_Exportado = Server.CreateObject("ADODB.RecordSet")
				RS_Exportado.Open SQL_Exportado, Conexao

				RS_ForMobile.Close
				Set RS_ForMobile = Nothing
			
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
		RS_Arquivos.Open SQL_Arquivos, Conexao_2012
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
        <big style="background-color:#FF0;"><a href="arquivos_2012/<%=Filename%>" target="_blank">Clique aqui * para salvar o arquivo.</a></big>&nbsp;* Botão direito > Salvar Como	<br>
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
<% 
Conexao.Close 
Conexao_2012.Close
%>