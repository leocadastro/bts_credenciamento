<%
Server.ScriptTimeout = 99999
If Session("admin_xml_logado") <> true Then
	Session("admin_xml_msg") = "Por favor logue-se novamente"
	response.Redirect("/admin/xml_meils/")	
End If

	dim eventos(2)
	eventos(0) = array(27,"logo_fsnordeste.gif","Food Service Nordeste 2011")
	eventos(1) = array(28,"logo_ftnordeste.gif","Tecnologia Nordeste 2011")
	eventos(2) = array(29,"logo_abfnordeste2011.gif","ABF Nordeste 2011")
'	eventos(3) = array(24,"logo_fispalcafe2011.gif","Fispal Caf&eacute; 2011")
'	eventos(4) = array(20,"logo_ft.jpg","Fispal Tecnologia 2011")
'	eventos(5) = array(22,"logo_abf.jpg","ABF Franchising Expo 2011")
%>
<!--#include virtual="/admin/inc/acentuacao.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>... Brazil Trade Shows .. Credenciamento - ADM</title>
<script language="javascript" src="/js/maskInput.js"></script>
</head>
<link href="/css/css.css" rel="stylesheet" type="text/css" />
<style type="text/css">
.cursor  { cursor:pointer; }
</style>
<body class="conteudo_home">
<table width="520" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="/images/logo_bts_email.jpg" width="300" height="52" /></td>
  </tr>
  <tr>
    <td class="bemvindo">&nbsp;</td>
  </tr>
  <tr>
    <td class="bemvindo">
      <p align="center">ADM - Exportar XML</p>
    </td>
  </tr>
  <tr valign="top">
    <td align="center">&nbsp;</td>
  </tr>
  <tr valign="top">
    <td height="70" align="center">
	<% For i = 0 To Ubound(eventos) %>
      <% If Cstr(eventos(i)(0)) = Session("admin_xml_evento") Then %>
      <table width="271" border="0" cellspacing="0" cellpadding="0" background="/admin/images/bts/fundo_bts_menu.gif" class="cursor" onClick="document.location='evento.asp?id=<%=eventos(i)(0)%>';">
        <tr>
          <td height="54" align="center"><img src="/images/<%=eventos(i)(1)%>" alt="<%=eventos(i)(2)%>" title="<%=eventos(i)(2)%>"></td>
        </tr>
      </table>
      <% End If %>
      <% Next %></td>
  </tr>
  <tr>
    <td align="center" class="conteudo_home">&nbsp;</td>
  </tr>
</table>
<table width="520" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="25%" align="center" class="conteudo_home"><a href="menu.asp">Menu</a></td>
    <td width="25%" align="center" class="conteudo_home"><a href="exportar.asp">Exportar Diferencial</a></td>
    <td width="25%" align="center" class="conteudo_home"><a href="arquivos.asp">Arquivos Gerados</a></td>
    <td width="25%" align="center" class="conteudo_home"><a href="logout.asp">Log Out</a></td>
  </tr>
</table>
<br>
<table width="520" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
	<td class="conteudo_home">
	<%
	cod_evento = Session("admin_xml_evento") 

	Set Conexao = Server.CreateObject("ADODB.Connection")
	Conexao.Provider = "SQLOLEDB"
	Conexao.Open Application("cnn")

	SQL_Cadastros = "SET TRANSACTION ISOLATION LEVEL REPEATABLE READ; " &_
					"BEGIN TRANSACTION;  " &_
					"Select Distinct top 1000 " &_
					"	PF.cod_cadastro_PF,  " &_
					" 	PJ.txt_cnpj as CNPJ, " &_
					"	PJ.txt_razao as Empresa,  " &_
					"	PJ.txt_fantasia as NomeFantasia, " &_
					"	PF.txt_nome as Nome,  " &_
					"	PF.txt_sobrenome as SobreNome, " &_
					"	PF.txt_nome_cred as Credencial, " &_
					"	PF.sexo as Sexo, " &_
					"	PF.txt_cpf as CPF,  " &_
					"	PF.dt_nascimento as Data_Nascimento, " &_
					"	PF.nu_ddi_tel as DDI_Telefone,   " &_
					"	Right(PF.nu_ddd_tel,2) as DDD_Telefone,   " &_
					"	PF.nu_telefone as Telefone,   " &_
					"	PF.nu_ddi_cel as DDI_Celular,   " &_
					"	Right(PF.nu_ddd_cel,2) as DDD_Celular,   " &_
					"	PF.nu_celular as Celular, " &_
					"	PF.in_SMS as SMS, " &_
					"	PF.txt_email as Email, " &_  
					"	PF.in_Email as [Opt-in], " &_
					"	CG.txt_cargo_br as Cargo,  " &_
					"	PF.txt_cargo_outros as Cargo_Outros,  " &_
					"	SC.txt_subcargo_br as SubCargo,  " &_
					"	DP.txt_depto_br as Depto,   " &_
					"	PF.txt_depto_outros as Depto_Outros,  " &_
					"	PJ.txt_endereco as Endereco,  " &_
					"	PJ.txt_numero as Numero,  " &_
					"	PJ.txt_complemento as Complemento,  " &_
					"	PJ.txt_bairro as Bairro, " &_
					"	PJ.txt_cidade as Cidade,   " &_
					"	PJ.txt_uf as UF,   " &_
					"	PJ.txt_pais as Pais, " &_
					"	PJ.txt_cep as CEP, " &_
					"	PJ.txt_site as Site, " &_
					"	RA.txt_ramo_br as Ramo,   " &_
					"	PJ.txt_ramoatividade as Ramo_Outros,   " &_
					"	PJ.txt_produto as Principal_Produto, " &_
					"	FU.txt_funcionarios_br as Qtde_Funcionarios, " &_
					"	PJ.in_afiliada_ent_classe as Entidade_Classe, " &_
					"	PJ.txt_afiliada_ent_classe as Nome_Entidade " &_
					"From CadastroPF PF (nolock) " &_
					"Left Join CadastroPJ PJ (nolock) on PJ.cod_cadastro_PJ = PF.cod_cadastro_PJ  " &_
					"Left Join Cargo CG (nolock) on CG.cod_cargo = PF.cod_cargo  " &_
					"Left Join SubCargo SC (nolock) on SC.cod_subcargo = PF.cod_subcargo  " &_
					"Left Join Depto DP (nolock) on DP.cod_depto = PF.cod_depto  " &_
					"Left Join RamoAtividade RA (nolock) on RA.cod_ramo = PJ.cod_ramo  " &_
					"Left Join Funcionarios as FU (nolock) on FU.cod_funcionarios = PJ.cod_funcionarios "&_ 
					"Where  " &_
					"	PF.cod_evento = " & cod_evento & " " &_
					"	AND exportado = 0  " &_
					"Order by PF.cod_cadastro_PF " &_
					"COMMIT TRANSACTION;  "

'response.write(SQL_Cadastros & "<hr>")

	Set RS_Cadastros = Server.CreateObject("ADODB.RecordSet")
	RS_Cadastros.CursorType = 0
	RS_Cadastros.LockType = 1
	RS_Cadastros.Open SQL_Cadastros, Conexao
	
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
	
	If not RS_Cadastros.BOF or not RS_Cadastros.EOF Then

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
		
'		If Session("admin_xml_evento") = "9" Then
'			feira = "FT2010"
'		ElseIf Session("admin_xml_evento") = "17" Then
'			feira = "FHR2010"
'		ElseIf Session("admin_xml_evento") = "13" Then
'			feira = "MA2010"
'		End If

		For i = 0 To Ubound(eventos)
			If Cstr(eventos(i)(0)) = Session("admin_xml_evento") Then
				feira = Replace(eventos(i)(2)," ", "_")
				Exit For
			End If
		Next
		
		extensao = ".xml"

		Filename = arquivo & "_" & feira & "_" & data & extensao ' file to read

		Const ForReading = 1, ForWriting = 2, ForAppending = 3
		Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

		' Create a filesystem object
		Dim FSO
		set FSO = server.createObject("Scripting.FileSystemObject")

		' Map the logical path to the physical system path
		Dim Filepath
		Filepath = Server.MapPath("exportados/" & Filename)

		Set oFiletxt = FSO.CreateTextFile(Filepath, True)
		sPath = FSO.GetAbsolutePathName(Filepath)
		sFilename = FSO.GetFileName(sPath)

		oFiletxt.WriteLine("<?xml version='1.0' encoding='iso-8859-1'?>")
		oFiletxt.WriteLine("<credenciamento>")

		total = 0
		If not RS_Cadastros.BOF or not RS_Cadastros.EOF Then
			While not RS_Cadastros.EOF			
				total = total + 1
				response.write(total & " - " & RS_Cadastros("cod_cadastro_PF") & " - " & Ucase(RS_Cadastros("cpf")) & " - " & Ucase(RS_Cadastros("nome")) & " " & Ucase(RS_Cadastros("SobreNome")) & "<br>")
				response.write("<script>self.scrollBy(0,400)</script>")
				Response.Flush
				oFiletxt.WriteLine("<cadastro>")
				oFiletxt.WriteLine("<cod_identificacao>" 	& trocar(Ucase(RS_Cadastros("cod_cadastro_PF"))) & "</cod_identificacao>")
				oFiletxt.WriteLine("<cnpj>" 				& trocar(Ucase(RS_Cadastros("cnpj"))) & "</cnpj>")
				oFiletxt.WriteLine("<empresa>" 				& trocar(Ucase(RS_Cadastros("empresa"))) & "</empresa>")
				oFiletxt.WriteLine("<nome_fantasia>"		& trocar(Ucase(RS_Cadastros("NomeFantasia"))) & "</nome_fantasia>")
				oFiletxt.WriteLine("<cpf>" 					& trocar(Ucase(RS_Cadastros("cpf"))) & "</cpf>")
				oFiletxt.WriteLine("<nome>" 				& trocar(Ucase(RS_Cadastros("nome"))) & " " & Ucase(RS_Cadastros("SobreNome")) & "</nome>")
				oFiletxt.WriteLine("<credencial>" 			& trocar(Ucase(RS_Cadastros("credencial"))) & "</credencial>")
				oFiletxt.WriteLine("<sexo>"					& trocar(Ucase(RS_Cadastros("sexo"))) & "</sexo>")
				oFiletxt.WriteLine("<data_nascimento>"		& trocar(Ucase(RS_Cadastros("data_nascimento"))) & "</data_nascimento>")
				oFiletxt.WriteLine("<ddi_telefone>" 		& trocar(Ucase(RS_Cadastros("ddi_telefone"))) & "</ddi_telefone>")
				oFiletxt.WriteLine("<ddd_telefone>" 		& trocar(Ucase(RS_Cadastros("ddd_telefone"))) & "</ddd_telefone>")
				oFiletxt.WriteLine("<telefone>"				& trocar(Ucase(RS_Cadastros("telefone"))) & "</telefone>")
				oFiletxt.WriteLine("<ddi_celular>"	 		& trocar(Ucase(RS_Cadastros("ddi_celular"))) & "</ddi_celular>")
				oFiletxt.WriteLine("<ddd_celular>"	 		& trocar(Ucase(RS_Cadastros("ddd_celular"))) & "</ddd_celular>")
				oFiletxt.WriteLine("<celular>"				& trocar(Ucase(RS_Cadastros("celular"))) & "</celular>")
				oFiletxt.WriteLine("<sms>" 					& trocar(Ucase(RS_Cadastros("sms"))) & "</sms>")
				oFiletxt.WriteLine("<email>" 				& trocar(Ucase(RS_Cadastros("email"))) & "</email>")
				oFiletxt.WriteLine("<opt_in>"				& trocar(Ucase(RS_Cadastros("opt-in"))) & "</opt_in>")
				oFiletxt.WriteLine("<cargo>"				& trocar(Ucase(RS_Cadastros("cargo"))) & "</cargo>")
				oFiletxt.WriteLine("<cargo_outros>" 		& trocar(Ucase(RS_Cadastros("cargo_outros"))) & "</cargo_outros>")
				oFiletxt.WriteLine("<subcargo>" 			& trocar(Ucase(RS_Cadastros("SubCargo"))) & "</subcargo>")
				oFiletxt.WriteLine("<departamento>" 		& trocar(Ucase(RS_Cadastros("depto"))) & "</departamento>")
				oFiletxt.WriteLine("<departamento_outros>" 	& trocar(Ucase(RS_Cadastros("depto_outros"))) & "</departamento_outros>")
				oFiletxt.WriteLine("<endereco>" 			& trocar(Ucase(RS_Cadastros("endereco"))) & "</endereco>")
				oFiletxt.WriteLine("<nro>" 					& trocar(Ucase(RS_Cadastros("numero"))) & "</nro>")
				oFiletxt.WriteLine("<complemento>" 			& trocar(Ucase(RS_Cadastros("complemento"))) & "</complemento>")
				oFiletxt.WriteLine("<bairro>" 				& trocar(Ucase(RS_Cadastros("bairro"))) & "</bairro>")
				oFiletxt.WriteLine("<cidade>" 				& trocar(Ucase(RS_Cadastros("cidade"))) & "</cidade>")
				oFiletxt.WriteLine("<uf>" 					& trocar(Ucase(RS_Cadastros("uf"))) & "</uf>")
				oFiletxt.WriteLine("<pais>" 				& trocar(Ucase(RS_Cadastros("pais"))) & "</pais>")
				oFiletxt.WriteLine("<cep>" 					& trocar(Ucase(RS_Cadastros("cep"))) & "</cep>")
				oFiletxt.WriteLine("<site>" 				& trocar(Ucase(RS_Cadastros("site"))) & "</site>")
				oFiletxt.WriteLine("<ramo_atividade>" 		& trocar(Ucase(RS_Cadastros("ramo"))) & "</ramo_atividade>")
				oFiletxt.WriteLine("<ramo_atividade_outros>" & trocar(Ucase(RS_Cadastros("ramo_outros"))) & "</ramo_atividade_outros>")
				oFiletxt.WriteLine("<principal_produto>"	& trocar(Ucase(RS_Cadastros("principal_produto"))) & "</principal_produto>")
				oFiletxt.WriteLine("<qtde_funcionarios>"	& trocar(Ucase(RS_Cadastros("qtde_funcionarios"))) & "</qtde_funcionarios>")
				oFiletxt.WriteLine("<entidade_classe>" 		& trocar(Ucase(RS_Cadastros("entidade_classe"))) & "</entidade_classe>")
				oFiletxt.WriteLine("<nome_entidade>"		& trocar(Ucase(RS_Cadastros("nome_entidade"))) & "</nome_entidade>")
				
				SQL_RespostasPesquisa = 	"Select " &_
											"	txt_pergunta, " &_
											"	txt_resposta " &_
											"From RespostasPesquisa " &_
											"Where cod_cadastro_PF = " & RS_Cadastros("cod_cadastro_PF") & " " &_
											"Order by txt_pergunta " 
											
'											"	Case txt_pergunta " &_
'											"		When 'Sua empresa deverá investir em 2009:' Then 'Sua empresa deverá investir em:' " &_
'											"		Else txt_pergunta " &_
'											"	End as txt_pergunta,  " &_
											
'	response.write("<hr>" & SQL_RespostasPesquisa & "<hr>")
											
				Set RS_RespostasPesquisa = Server.CreateObject("ADODB.RecordSet")
				RS_RespostasPesquisa.CursorType = 0
				RS_RespostasPesquisa.LockType = 1
				RS_RespostasPesquisa.Open SQL_RespostasPesquisa, Conexao
				
				If not RS_RespostasPesquisa.BOF or not RS_RespostasPesquisa.EOF Then
					oFiletxt.WriteLine("<pesquisa>")
					While not RS_RespostasPesquisa.EOF
						pergunta = Replace(RS_RespostasPesquisa("txt_pergunta"), "2009", Year(Now()))
						resposta = RS_RespostasPesquisa("txt_resposta")
						If Len(pergunta) > 0 Then
							oFiletxt.WriteLine("<pergunta questao='" & trocar(pergunta) & "' resposta='" & trocar(resposta) & "'/>")
						End If
						RS_RespostasPesquisa.MoveNext
					Wend
					oFiletxt.WriteLine("</pesquisa>")
					RS_RespostasPesquisa.Close
				End If				
				oFiletxt.WriteLine("</cadastro>")

				SQL_Exportado = "Update CadastroPF " &_
								"Set exportado = 1, " &_
								"dt_exportado = getDate() " &_
								"Where cod_cadastro_PF = " & RS_Cadastros("cod_cadastro_PF")
				Set RS_Exportado = Server.CreateObject("ADODB.RecordSet")
				RS_Exportado.Open SQL_Exportado, Conexao

				RS_Cadastros.MoveNext()
				response.Flush()
			Wend
			RS_Cadastros.Close
		End If

		oFiletxt.WriteLine("</credenciamento>")
		oFiletxt.Close

		SQL_Arquivos = 	"Insert Into Arquivos_XML " &_
						"(arquivo, total, cod_evento) " &_
						"values " &_
						"('" & filename & "'," & total & "," & cod_evento & ")"
		Set RS_Arquivos = Server.CreateObject("ADODB.RecordSet")
		RS_Arquivos.Open SQL_Arquivos, Conexao


		%>
        <hr>
		Arquivo <B><%=Filename%></B> criado com sucesso<br><br>
		Total de Cadastros Listados : <b><%=total%></b><br>
		<a href="/admin/xml_meils/exportados/<%=Filename%>" target="_blank">Clique aqui * para salvar o arquivo</a><br>
        * Botão direito > Salvar Como
		<%
		response.write("<script>self.scrollBy(0,400)</script>")
		Response.Flush
		%>
	<% Else %>
    	Não existem novos cadastros.
    <% End IF %>

	</td>
  </tr>
</table>
</body>
</html>
<% Conexao.Close %>