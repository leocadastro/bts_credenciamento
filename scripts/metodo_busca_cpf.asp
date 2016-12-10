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
	If RS_Verificar.BOF or RS_Verificar.EOF Then
	
		'=======================================================================
		Set ConexaoCredOld 	= Server.CreateObject("ADODB.Connection")
							  ConexaoCredOld.Open Application("cnnCredOld")
		'=======================================================================			
			
			'=======================================================================
			' Verificando o cadastro de Pessoa Fisica
			'=======================================================================
			SQL_Verifica = "SELECT " &_ 
							"	txt_nome 			as Nome " &_ 
							"	,txt_sobrenome 		as Sobrenome " &_ 
							"	,txt_nome_cred 		as NomeCredencial " &_ 
							"	,dt_nascimento 		as DtNasc " &_ 
							"	,nu_ddi_tel 		as Tel1DDI " &_ 
							"	,nu_ddd_tel 		as Tel1DDD " &_ 
							"	,nu_telefone 		as Tel1Fone " &_ 
							"	,nu_ddi_cel 		as Tel2DDI " &_ 
							"	,nu_ddd_cel 		as Tel2DDD " &_ 
							"	,nu_celular 		as Tel2Fone " &_
							"	,in_SMS 			as SMS " &_ 
							"	,cod_cargo 			as Cargo " &_ 
							"	,txt_cargo_outros 	as CargoOutros " &_ 
							"	,cod_subcargo 		as SubCargo " &_ 
							"	,cod_depto 			as Depto " &_ 
							"	,txt_depto_outros 	as DeptoOutros " &_ 
							"	,txt_email 			as Email " &_ 
							"	,in_Email 			as inEmail " &_ 
							"	,sexo 				as Sexo " &_ 
							"FROM CadastroPF " &_
							"WHERE txt_cpf = '" & CPFOld & "' " &_ 
							"ORDER BY cod_cadastro_PF DESC"

			'response.Write("<strong>SQL_Verifica</strong><hr>" & SQL_Verifica & "<hr>")

			Set RS_Verificar = Server.CreateObject("ADODB.Recordset")
			RS_Verificar.Open SQL_Verifica, ConexaoCredOld
			
			If RS_Verificar.BOF or RS_Verificar.EOF Then
		
				%>
                { 
                Resultado : '0',
                ResultadoTXT : 'Cadastro não localizado - CPF'
                }
                <%
		
			Else
			
				'=======================================================================
				' Cadastro localizado no Banco Antigo de PF
				' TRATANDO AS VARIAVEIS
				'=======================================================================
									
				NomeF			= Trim(RS_Verificar("Nome")) & " " & Trim(RS_Verificar("Sobrenome"))

				NomeCredF		= Trim(RS_Verificar("NomeCredencial"))
				
				If Len(Trim(NomeCredF)) > 27 Then NomeCredF = ""
				
				' Tratando a data de Nascimento
				If Len(Trim(RS_Verificar("DtNasc"))) = 8 Then
					SeparaData	= Split(Trim(RS_Verificar("DtNasc")), "/")
					DtNasc		= SeparaData(0) & "/" & SeparaData(1) & "/19" & SeparaData(2)
				Else
					DtNasc = Trim(RS_Verificar("DtNasc"))
				End if
							
				Tel1DDI			= Trim(RS_Verificar("Tel1DDI"))
				Tel1DDD			= Trim(RS_Verificar("Tel1DDD"))
				Tel1Fone		= Trim(RS_Verificar("Tel1Fone"))
				Tel1Fone		= Replace(Tel1Fone,"-","")
				
				' Tratando telefone 1
				' If Len(Trim(RS_Verificar("Tel1Fone"))) = 8 Then
				'	Tel			= Trim(RS_Verificar("Tel1Fone"))
				'	Tel1		= Left(tel,4)
				'	Tel2		= Right(tel,4)
				'	Tel1Fone	= Tel1 & "-" & Tel2
				'Else
				'	Tel1Fone	= Trim(RS_Verificar("Tel1Fone"))
				'End if
				
				Tel2DDI			= Trim(RS_Verificar("Tel2DDI"))
				Tel2DDD			= Trim(RS_Verificar("Tel2DDD"))
				Tel2Fone		= Trim(RS_Verificar("Tel2Fone"))
				Tel2Fone		= Replace(Tel2Fone,"-","")
				
				' Tratando Cargo
				If Len(RS_Verificar("Cargo")) > 0 Then
					CargoF		= Trim(RS_Verificar("Cargo"))
				Else	
					CargoF		= Trim(RS_Verificar("CargoOutros"))
				End If
				
				' Tratando Depto
				If Len(RS_Verificar("Depto")) > 0 Then
					DeptoF		= Trim(RS_Verificar("Depto"))
				Else
					DeptoF		= Trim(RS_Verificar("DeptoOutros"))
				End If	
				
				Email			= Trim(Lcase(RS_Verificar("Email")))
				Sexo			= Trim(RS_Verificar("Sexo"))
				
				'=======================================================================
				' Retornando JSON
				'=======================================================================
				
				%>
				{ 
				"NomeF" : "<%=NomeF%>",
				"NomeCredencialF" : "<%=NomeCredF%>",
				"DTNasc" : "<%=DtNasc%>",
				"Cargo" : "<%=CargoF%>",
				"Departamento" : "<%=DeptoF%>",
				"DDI1" : "<%=Tel1DDI%>",
				"DDD1" : "<%=Tel1DDD%>",
				"Fone1" : "<%=Tel1Fone%>",
				"Ramal1" : "",
				"DDI2" : "<%=Tel2DDI%>",
				"DDD2" : "<%=Tel2DDD%>",
				"Fone2" : "<%=Tel2Fone%>",
				"Ramal2" : "",
				"Email" : "<%=Email%>",
				"Sexo" : "<%=Sexo%>",
				"Banco" : "Old", 
				"Resultado" : "1",
				"ResultadoTXT" : "Cadastro localizado"
				}
				<%
						
			End If
			
	Else
	
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
		

	End If	
	
Else 

'=======================================================================
' Cadastro NÃO localizado 
'=======================================================================


End If

Conexao.Close
%>