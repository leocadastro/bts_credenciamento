<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%> 
<%
Response.Expires = -1
Response.AddHeader "Cache-Control", "no-cache"
Response.AddHeader "Pragma", "no-cache" 
%>
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/acentos_htm.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<%
response.Charset = "utf-8" 
response.ContentType = "text/html" 

If Session("cliente_edicao") = "" OR Session("cliente_idioma") = "" OR Session("cliente_tipo") = "" Then
	response.Redirect("/?erro=1")
End If

'=======================================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
			  Conexao.Open Application("cnn")
'=======================================================================

'=======================================================================
' Pegando o CPF para comparar com o Banco novo e com o Banco Antigo
'=======================================================================
CPFOld = Trim(Limpar_Texto(Request("cpf")))
CPF 	= Trim(Limpar_Texto(Request("cpf")))
CPF 	= Replace(CPF,".","")
CPF 	= Replace(CPF,"-","")

If Session("cliente_idioma") <> "1" Then
	Documento = "Passaporte"
Else 
	Documento = "CPF"
End If

'=======================================================================
' Verificando valor do campo documento
'=======================================================================
If Len(CPF) <> 0 Then

	'=======================================================================
	' Verificando cadastro no banco novo Credenciamento_2012
	'=======================================================================
	SQL_Verificar =	"SELECT top 1 " &_
					"	ID_Visitante " &_
					"	,Nome_Completo		as Nome " &_
					"	,Nome_Credencial	as NomeCredencial " &_
					"	,Data_Nasc			as DtNasc " &_
					"	,Sexo				as Sexo " &_
					"	,Email				as Email " &_
					"	,ID_Cargo " &_
					"	,Cargo_Outros " &_
					"	,ID_Depto " &_
					"	,Depto_Outros " &_
					"FROM Visitantes " &_ 
					"WHERE " & Documento & " = '" & CPF & "' " &_
					"	Order by ID_VISITANTE DESC "
	'response.Write("<strong>SQL_Verificar</strong><hr>" & SQL_Verificar & "<hr>")
	Set RS_Verificar = Server.CreateObject("ADODB.Recordset")
	RS_Verificar.Open SQL_Verificar, Conexao

	'=======================================================================
	' Verificando o Retorno da Query SQL_Verificar
	'=======================================================================
	If Not RS_Verificar.EOF Then
		'=======================================================================
		' Cadastro localizado no Banco Novo Credenciamento_2012
		' TRATANDO AS VARIAVEIS
		'=======================================================================
		ID_Visitante	= RS_Verificar("ID_Visitante")
		
		NomeF			= RS_Verificar("Nome")
		NomeCredF		= Trim(RS_Verificar("NomeCredencial"))
		
		' Tratando a data de Nascimento
		If Len(Trim(RS_Verificar("DtNasc"))) > 8 Then
			SeparaData	= Split(Trim(RS_Verificar("DtNasc")), "/")
			DtNasc		= SeparaData(0) & "/" & SeparaData(1) & "/19" & SeparaData(2)
		Else
			
			SeparaData	 = Trim(RS_Verificar("DtNasc"))

			SeparaDataDia = Left(SeparaData,2)
			SeparaDataMes = Mid(SeparaData,3,2)
			SeparaDataAno = Right(SeparaData,4)

			DtNasc		= SeparaDataDia & "/" & SeparaDataMes & "/" & SeparaDataAno
		End if
		
		' Verificando telefones cadastrados
		SQL_Verificar_Telefone = 	"SELECT " &_
									"	ID_Tipo_Telefone " &_
									"	,DDI " &_ 
									"	,DDD " &_
									"	,Numero " &_
									"	,Ramal " &_
									"	,SMS " &_
									"FROM Relacionamento_Telefones " &_
									"WHERE " &_
									"	ID_Visitante = " & ID_Visitante &_
									"	AND Ativo =  1"

		'response.write("<hr>SQL_Verificar_Telefone:<hr>" & SQL_Verificar_Telefone & "<hr>")

		Set RS_Verificar_Telefone = Server.CreateObject("ADODB.Recordset")
		RS_Verificar_Telefone.Open SQL_Verificar_Telefone, Conexao	
	
		t = 1
		Telefones = ""

		While not RS_Verificar_Telefone.EOF

			Telefones = Telefones & """DDI" & t & """ : """ & Trim(RS_Verificar_Telefone("DDI")) & ""","
			Telefones = Telefones & """DDD" & t & """ : """ & Trim(RS_Verificar_Telefone("DDD")) & ""","
			Telefones = Telefones & """Fone" & t & """ : """ & Trim(RS_Verificar_Telefone("Numero")) & ""","
			Telefones = Telefones & """ID_Tipo_Telefone" & t & """ : """ & Trim(RS_Verificar_Telefone("ID_Tipo_Telefone")) & ""","
			Telefones = Telefones & """Ramal" & t & """ : """ & Trim(RS_Verificar_Telefone("Ramal")) & ""","
			Telefones = Telefones & """SMS" & t & """ : """ & Trim(RS_Verificar_Telefone("SMS")) & ""","
			t = t + 1

			RS_Verificar_Telefone.MoveNext()
		Wend
		RS_Verificar_Telefone.Close	
		
		' Tratando Cargo
		CargoF			= Trim(RS_Verificar("ID_Cargo"))
		CargoOutros		= Trim(RS_Verificar("Cargo_Outros"))
				
		DeptoF			= Trim(RS_Verificar("ID_Depto"))
		DeptoOutros		= Trim(RS_Verificar("Depto_Outros"))
		
		Email			= Trim(Lcase(RS_Verificar("Email")))
		Sexo			= Trim(RS_Verificar("Sexo"))
		
		' Verificar se já preencheu na Edição da SESSÃO
		SQL_Confirmar_mesma_edicao = 	"Select " &_
										"	ID_Relacionamento_Cadastro " &_
										"From Relacionamento_Cadastro " &_
										"Where 	ID_Visitante = 	" & ID_Visitante & " " &_
										"		AND ID_Edicao = " & Session("cliente_edicao")
		Set RS_Confirmar_mesma_edicao = Server.CreateObject("ADODB.Recordset")
		RS_Confirmar_mesma_edicao.Open SQL_Confirmar_mesma_edicao, Conexao	
		
		' se NAO EXISTE, continua
		If RS_Confirmar_mesma_edicao.BOF or RS_Confirmar_mesma_edicao.EOF Then
		
			'=======================================================================
			' Retornando JSON
			'=======================================================================
			
			%>
			{ 
			"ID_Visitante" : "<%=ID_Visitante%>",
			"NomeF" : "<%=NomeF%>",
			"NomeCredencialF" : "<%=NomeCredF%>",
			"DTNasc" : "<%=DtNasc%>",
			"Cargo" : "<%=CargoF%>",
			"CargoOutros" : "<%=CargoOutros%>",
			"Departamento" : "<%=DeptoF%>",
			"DepartamentoOutros" : "<%=DeptoOutros%>",
			<%=Telefones%>
			"Email" : "<%=Email%>",
			"Sexo" : "<%=Sexo%>",
			"Banco" : "New", 
			"Resultado" : "1",
			"ResultadoTXT" : "Cadastro localizado"
			}
			<%

		Else
			'=======================================================================
			' Retornando JSON
			'=======================================================================
			%>
			{ 
			Resultado : '2',
			ResultadoTXT : 'Cadastro já realizado nesta Edição - CPF'
			}
			<%
		End If
	Else
	
	%>
	{ 
	Resultado : '0',
	ResultadoTXT : 'Cadastro não localizado - CPF'
	}
	<%
		

	End If	
	
Else 

'=======================================================================
' Cadastro NÃO localizado 
'=======================================================================


End If

Conexao.Close
%>