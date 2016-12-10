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
' Pegando o CNPJ para comparar com o Banco novo e com o Banco Antigo
'=======================================================================
CNPJOld = Trim(Limpar_Texto(Request("cnpj")))
CNPJ 	= Trim(Limpar_Texto(Request("cnpj")))
CNPJ 	= Replace(CNPJ,".","")
CNPJ 	= Replace(CNPJ,"-","")
CNPJ 	= Replace(CNPJ,"/","")

'=======================================================================
' Verificando valor do campo documento
'=======================================================================
If Len(CNPJ) <> 0 Then

	'=======================================================================
	' Verificando cadastro no banco novo Credenciamento_2012
	'=======================================================================
	SQL_Cred2012_Empresa =	"SELECT ID_Empresa " &_ 
							"	,ID_Formulario " &_ 
							"	,ID_Funcionarios_Qtde " &_ 
							"	,cnpj 					as CNPJ " &_ 
							"	,Razao_Social 			as Razao " &_ 
							"	,Nome_Fantasia			as Fantasia " &_ 
							"	,Presidente				as Sigla " &_ 
							"	,Principal_Produto		as Produto " &_ 
							"	,Site 					as Site " &_ 
							"	,Email 					as Email " &_ 
							"	,Newsletter 			as Newsletter " &_ 
							"	,Reitor			 		as Coordenador " &_ 
							"	,Senha 					as Senha " &_ 
							"	,Ativo					as Ativo " &_ 
							"FROM Empresas " &_ 
							"WHERE CNPJ = '" & CNPJ & "' " &_
							"	AND Ativo = 1 " &_
							"Order by ID_Empresa DESC"
	'response.Write("<strong>SQL_Verificar</strong><hr>" & SQL_Cred2012_Empresa & "<hr>")
	Set RS_Cred2012_Empresa = Server.CreateObject("ADODB.Recordset")
	RS_Cred2012_Empresa.CursorType = 0
	RS_Cred2012_Empresa.LockType = 1
	RS_Cred2012_Empresa.Open SQL_Cred2012_Empresa, Conexao

	'=======================================================================
	' Verificando o Retorno da Query SQL_Verificar - Credenciamento_2012
	'=======================================================================
	If Not RS_Cred2012_Empresa.EOF Then
		' Nao existe procurar no Banco ANTIGO
		'=======================================================================
		Set ConexaoCredOld 	= Server.CreateObject("ADODB.Connection")
							  ConexaoCredOld.Open Application("cnnCredOld")
		'=======================================================================			
		
		'=======================================================================
		' Verificando cadastro no banco antigo CredenciamentoBTS
		'=======================================================================
		
		'=======================================================================
		' A) Verificar banco de Universidades Prineiro
		'=======================================================================
		SQL_Universidade = 	"SELECT top 1 " &_
							"	Assoc.cod_cadastro_Assoc		as CodAssoc  " &_
							"	,Assoc.cod_cadastro_PF			as CodPF  " &_
							"	,Assoc.cod_evento 				as Evento  " &_
							"	,Assoc.cod_idioma 				as Idioma  " &_
							"	,Assoc.txt_cnpj 				as CNPJ " &_
							"	,Assoc.txt_razao 				as Razao  " &_
							"	,Assoc.txt_fantasia 			as Fantasia  " &_
							"	,Assoc.txt_sigla	 			as Sigla  " &_
							"	,Assoc.txt_cep 					as CEP  " &_
							"	,Assoc.txt_endereco 			as Endereco  " &_
							"	,Assoc.txt_numero 				as Numero  " &_
							"	,Assoc.txt_complemento 			as Complemento  " &_
							"	,Assoc.txt_bairro 				as Bairro  " &_
							"	,Assoc.txt_cidade 				as Cidade  " &_
							"	,Assoc.txt_uf 					as UF  " &_
							"	,Assoc.txt_pais 				as Pais  " &_
							"	,Assoc.txt_site 				as Site  " &_
							"	,Assoc.dt_cadastro 				as DataCad  " &_
							"FROM CadastroAssoc as Assoc " &_
							"INNER JOIN CadastroPF as PF " &_
							"	ON Assoc.cod_cadastro_PF = PF.cod_cadastro_PF " &_
							"WHERE " &_
							"	txt_cnpj = '" & CNPJOld & "' " &_
							"	AND txt_nome_presidente = ''	/* ==>	Garante que é UNIVERSIDADE, Campo obrigatório de ENTIDADE	<==	*/ " &_
							"ORDER BY Assoc.cod_cadastro_Assoc DESC, Assoc.dt_cadastro DESC "
    
		'response.Write("<strong>SQL_Universidade</strong><hr>" & SQL_Universidade & "<hr>")
    
		Set RS_Universidade = Server.CreateObject("ADODB.Recordset")
		RS_Universidade.CursorType = 0
		RS_Universidade.LockType = 1
		RS_Universidade.Open SQL_Universidade, ConexaoCredOld
    
		If not RS_Universidade.BOF or not RS_Universidade.EOF Then
		
			CNPJ			= Trim(RS_Universidade("CNPJ"))
			CNPJ 			= Replace(CNPJ,".","")
			CNPJ 			= Replace(CNPJ,"-","")
			CNPJ 			= Replace(CNPJ,"/","")
			Razao			= Trim(RS_Universidade("Razao"))
			Fantasia		= Trim(RS_Universidade("Fantasia"))
			Sigla			= Trim(RS_Universidade("Sigla"))
			Produto			= ""
			CEP				= Trim(RS_Universidade("CEP"))
			Endereco		= Trim(RS_Universidade("Endereco"))
			Numero			= Trim(RS_Universidade("Numero"))
			Complemento		= Trim(RS_Universidade("Complemento"))
			Bairro			= Trim(RS_Universidade("Bairro"))
			Cidade			= Trim(RS_Universidade("Cidade"))
			UF				= Trim(RS_Universidade("UF"))
			Pais			= Trim(RS_Universidade("Pais"))
			
			Site			= Replace(RS_Universidade("Site"),"http://","")
			Site			= Trim(Lcase(Site))
			
			%>
				{ 
				Razao : '<%=Razao%>',
				Fantasia : '<%=Fantasia%>',
				Sigla : '<%=Sigla%>',
				Produto : '<%=Produto%>',
				CEP : '<%=CEP%>',
				Endereco : '<%=Endereco%>',
				Numero : '<%=Numero%>',
				Complemento : '<%=Complemento%>',
				Bairro : '<%=Bairro%>',
				Cidade : '<%=Cidade%>',
				UF : '<%=UF%>',
				Pais : '<%=Pais%>',
				Site : '<%=Site%>',
   				Banco : 'Empresa_OLD', 
				Resultado : '1',
				ResultadoTXT : 'Cadastro localizado'
                }
                <%
		
		Else
    
			SQL_Empresa =		"SELECT top 1 " &_
								"	PJ.cod_cadastro_PJ 			as CodPJ " &_
								"	,PJ.cod_cadastro_PF			as CodPF " &_
								"	,PJ.cod_evento 				as Evento " &_
								"	,PJ.cod_idioma 				as Idioma " &_
								"	,PJ.txt_cnpj 				as CNPJ " &_
								"	,PJ.txt_razao 				as Razao " &_
								"	,PJ.txt_fantasia 			as Fantasia " &_
								"	,PJ.txt_cep 				as CEP " &_
								"	,PJ.txt_endereco 			as Endereco " &_
								"	,PJ.txt_numero 				as Numero " &_
								"	,PJ.txt_complemento 		as Complemento " &_
								"	,PJ.txt_bairro 				as Bairro " &_
								"	,PJ.txt_cidade 				as Cidade " &_
								"	,PJ.txt_uf 					as UF " &_
								"	,PJ.txt_pais 				as Pais " &_
								"	,PJ.txt_site 				as Site " &_
								"	,PJ.cod_ramo 				as IDRamo" &_
								"	,PJ.txt_ramoatividade 		as Ramo " &_
								"	,PJ.txt_produto 			as Produto " &_
								"	,PJ.cod_funcionarios 		as CodFunc " &_
								"	,PJ.in_afiliada_ent_classe 	as Afiliada " &_
								"	,PJ.txt_afiliada_ent_classe	as AfiliadaTXT " &_
								"	,PJ.dt_cadastro 			as DataCad " &_
								"	,PF.txt_cpf	 				as CPF " &_
								"FROM CadastroPJ as PJ " &_
								"INNER JOIN CadastroPF as PF " &_
								"	ON PJ.cod_cadastro_PF = PF.cod_cadastro_PF " &_
								"WHERE txt_cnpj = '" & CNPJOld & "' " &_
								"ORDER BY PJ.cod_cadastro_PJ DESC, PJ.dt_cadastro DESC"
	
			'response.Write("<strong>SQL_Verificar</strong><hr>" & SQL_Verificar & "<hr>")
	
			Set RS_Empresa = Server.CreateObject("ADODB.Recordset")
			RS_Empresa.CursorType = 0
			RS_Empresa.LockType = 1
			RS_Empresa.Open SQL_Empresa, ConexaoCredOld
			
			'=======================================================================
			' Verificando o Retorno da Query SQL_Verificar
			'=======================================================================
	
			If RS_Empresa.BOF or RS_Empresa.EOF Then
			
				%>
				{ 
				Resultado : '0',
				ResultadoTXT : 'Cadastro não localizado - CNPJ'
				}
				<%
			
			Else
			
				CNPJ			= Trim(RS_Empresa("CNPJ"))
				CNPJ 			= Replace(CNPJ,".","")
				CNPJ 			= Replace(CNPJ,"-","")
				CNPJ 			= Replace(CNPJ,"/","")
				Razao			= Trim(RS_Empresa("Razao"))
				Fantasia		= Trim(RS_Empresa("Fantasia"))
				Produto			= Trim(RS_Empresa("Produto"))
				CEP				= Trim(RS_Empresa("CEP"))
				Endereco		= Trim(RS_Empresa("Endereco"))
				Numero			= Trim(RS_Empresa("Numero"))
				Complemento		= Trim(RS_Empresa("Complemento"))
				Bairro			= Trim(RS_Empresa("Bairro"))
				Cidade			= Trim(RS_Empresa("Cidade"))
				UF				= Trim(RS_Empresa("UF"))
				Pais			= Trim(RS_Empresa("Pais"))
				
				Site			= Replace(RS_Empresa("Site"),"http://","")
				Site			= Trim(Lcase(Site))
				
				%>
					{ 
					Razao : '<%=Razao%>',
					Fantasia : '<%=Fantasia%>',
					Sigla : '',
					Produto : '<%=Produto%>',
					CEP : '<%=CEP%>',
					Endereco : '<%=Endereco%>',
					Numero : '<%=Numero%>',
					Complemento : '<%=Complemento%>',
					Bairro : '<%=Bairro%>',
					Cidade : '<%=Cidade%>',
					UF : '<%=UF%>',
					Pais : '<%=Pais%>',
					Site : '<%=Site%>',
					Banco : 'Empresa_OLD', 
					Resultado : '1',
					ResultadoTXT : 'Cadastro localizado'
					}
					<%
			
			End If
		End If
		
	Else

		'=======================================================================
		' Cadastro localizado no Banco Novo Credenciamento_2012
		' TRATANDO AS VARIAVEIS
		'=======================================================================
		
		ID_Empresa		= Trim(RS_Cred2012_Empresa("ID_Empresa"))
		ID_Formulario	= Trim(RS_Cred2012_Empresa("ID_Formulario"))
		CNPJ			= Trim(RS_Cred2012_Empresa("CNPJ"))
		Razao			= Trim(RS_Cred2012_Empresa("Razao"))
		Fantasia		= Trim(RS_Cred2012_Empresa("Fantasia"))
		Sigla			= Trim(RS_Cred2012_Empresa("Sigla"))
		Produto			= Trim(RS_Cred2012_Empresa("Produto"))
		Site			= Replace(RS_Cred2012_Empresa("Site"),"http://","")
		Site			= Trim(Lcase(Site))

		' Verificando Endereco cadastrado
		SQL_Verificar_Endereco = 	"SELECT " &_ 
									"	CEP " &_ 
									"	,Endereco " &_ 
									"	,Numero " &_ 
									"	,Complemento " &_ 
									"	,Bairro " &_ 
									"	,Cidade " &_ 
									"	,ID_UF " &_ 
									"	,ID_Pais " &_ 
									"FROM Relacionamento_Enderecos " &_
									"WHERE " &_
									"	ID_Empresa = " & RS_Cred2012_Empresa("ID_Empresa") & " " &_
									"	AND Ativo =  1"
		'response.write("<hr>SQL_Verificar_Endereco:<hr>" & SQL_Verificar_Endereco & "<hr>")
		Set RS_Cred2012_Empresa_Endereco = Server.CreateObject("ADODB.Recordset")
		RS_Cred2012_Empresa_Endereco.Open SQL_Verificar_Endereco, Conexao	

		TratarCEP		= Trim(RS_Cred2012_Empresa_Endereco("CEP"))
		CEP1 = Left(TratarCEP,5)
		CEP2 = Right(TratarCEP,3)
		CEP = CEP1 & "-" & CEP2

		Endereco		= Trim(RS_Cred2012_Empresa_Endereco("Endereco"))
		Numero			= Trim(RS_Cred2012_Empresa_Endereco("Numero"))
		Complemento		= Trim(RS_Cred2012_Empresa_Endereco("Complemento"))
		Bairro			= Trim(RS_Cred2012_Empresa_Endereco("Bairro"))
		Cidade			= Trim(RS_Cred2012_Empresa_Endereco("Cidade"))
		ID_UF			= Trim(RS_Cred2012_Empresa_Endereco("ID_UF"))
		ID_Pais			= Trim(RS_Cred2012_Empresa_Endereco("ID_Pais"))

		'=======================================================================
		' Select de Estados
		SQL_Estado = 		"SELECT " &_
							"	Sigla " &_ 
							"FROM UF " &_
							"WHERE " &_
							"	ID_UF = " & ID_UF & " " &_
							"	AND Ativo = 1 "
		Set RS_Estado = Server.CreateObject("ADODB.Recordset")
		RS_Estado.CursorType = 0
		RS_Estado.LockType = 1
		RS_Estado.Open SQL_Estado, Conexao
		UF = RS_Estado("Sigla")
		
		'=======================================================================
		' Select de Paises
		SQL_Pais =  		"SELECT " &_
							"	Pais_PTB " &_
							"FROM Pais " &_
							"WHERE " &_
							"	ID_Pais = " & ID_Pais & " " &_
							"	AND Ativo = 1 "
		Set RS_Pais = Server.CreateObject("ADODB.Recordset")
		RS_Pais.CursorType = 0
		RS_Pais.LockType = 1
		RS_Pais.Open SQL_Pais, Conexao
		Pais = Ucase(RS_Pais("Pais_PTB"))

		'=======================================================================
		' Verificando telefones cadastrados
		SQL_Verificar_Telefone = 	"SELECT * " &_ 
									"FROM Relacionamento_Telefones " &_
									"WHERE " &_
									"	Ativo =  1 " &_
									"AND " &_
									"	ID_Empresa = " & RS_Cred2012_Empresa("ID_Empresa")
		'response.write("<hr>SQL_Verificar_Telefone:<hr>" & SQL_Verificar_Telefone & "<hr>")
		Set RS_Cred2012_Empresa_Telefone = Server.CreateObject("ADODB.Recordset")
		RS_Cred2012_Empresa_Telefone.Open SQL_Verificar_Telefone, Conexao	
	
		t = 1
		Telefones = ""

		While not RS_Cred2012_Empresa_Telefone.EOF
			
			Telefones = Telefones & """DDI" & t & """ : """ & Trim(RS_Cred2012_Empresa_Telefone("DDI")) & ""","
			Telefones = Telefones & """DDD" & t & """ : """ & Trim(RS_Cred2012_Empresa_Telefone("DDD")) & ""","
			Telefones = Telefones & """Fone" & t & """ : """ & Trim(RS_Cred2012_Empresa_Telefone("Numero")) & ""","

			t = t + 1

			RS_Cred2012_Empresa_Telefone.MoveNext()
		Wend
		RS_Cred2012_Empresa_Telefone.Close	
		
		'Email			= Trim(Lcase(RS_Verificar("Email")))
		'Sexo			= Trim(RS_Verificar("Sexo"))
		
		'=======================================================================
		' Buscando Ramos "previamente" cadastrados
		SQL_Cred2012_Ramos_V2 = 	"Select " &_
									"	RA.Ramo_Atv_PTB as Ramo, " &_
									"	REVRE.Complemento " &_
									"From Relacionamento_Empresa_Visitante_RamoAtv_Edicao_V2 as REVRE " &_
									"Inner Join Relacionamento_RamoeAtividade_Edicoes_V2 as RRE " &_
									"	ON	( " &_
									"		RRE.ID_Ramo_Atividade = REVRE.ID_Ramo_Atividade " &_
									"		AND RRE.ID_Edicao = REVRE.ID_Edicao " &_
									"		) " &_
									"Inner Join RamoeAtividade_V2 as RA " &_
									"	ON	( " &_
									"		RA.ID_Ramo_Atividade = RRE.ID_Ramo_Atividade " &_
									"		AND RA.ID_Ramo_Atividade = REVRE.ID_Ramo_Atividade " &_
									"		) " &_
									"Inner Join Empresas as E " &_
									"	ON " &_
									"		E.ID_Empresa = REVRE.ID_Empresa " &_
									"Where " &_
									"	REVRE.Ativo = 1 " &_
									"	AND REVRE.ID_Edicao = " & Session("cliente_edicao") & " " &_
									"	AND E.ID_Empresa = " & ID_Empresa & " " &_
									"Order by Ramo, Complemento "
		'response.write("<hr>SQL_Cred2012_Ramos_V2:<hr>" & SQL_Cred2012_Ramos_V2 & "<hr>")
		Set RS_Cred2012_Ramos_V2 = Server.CreateObject("ADODB.Recordset")
		RS_Cred2012_Ramos_V2.Open SQL_Cred2012_Ramos_V2, Conexao

		If not RS_Cred2012_Ramos_V2.BOF or not RS_Cred2012_Ramos_V2.EOF Then
			Ramos 		= ""
			Ramos 	= Ramos & """Ramos"" : [ "

			While not RS_Cred2012_Ramos_V2.EOF
				
				Ramos 	= Ramos & "{ "
				Ramos 	= Ramos & """Ramo"" : """ & Trim(RS_Cred2012_Ramos_V2("Ramo")) & ""","
				Ramos 	= Ramos & """Complemento"" : """ & Trim(RS_Cred2012_Ramos_V2("Complemento")) & """"
				Ramos 	= Ramos & "} "

				RS_Cred2012_Ramos_V2.MoveNext()
				If not RS_Cred2012_Ramos_V2.EOF Then
					Ramos = Ramos & ","
				End If
			Wend
			RS_Cred2012_Ramos_V2.Close

			Ramos 	= Ramos & "],"
		End If

		'=======================================================================
		' Buscando Produtos "previamente" cadastrados
		SQL_Cred2012_Produtos_V2 = 	"Select " &_
									"	Principal_Produto as Produto " &_
									"FROM " &_
									"	Relacionamento_Produto_Edicao_Empresa_Visitante_v2 " &_
									"Where " &_
									"	ID_Empresa = " & ID_Empresa & " " &_
									"Order by Principal_Produto "
		'response.write("<hr>SQL_Cred2012_Produtos_V2:<hr>" & SQL_Cred2012_Produtos_V2 & "<hr>")
		Set RS_Cred2012_Produtos_V2 = Server.CreateObject("ADODB.Recordset")
		RS_Cred2012_Produtos_V2.Open SQL_Cred2012_Produtos_V2, Conexao

		If not RS_Cred2012_Produtos_V2.BOF or not RS_Cred2012_Produtos_V2.EOF Then
			Produtos	= ""
			Produtos 	= Produtos & """Produtos"" : [ "

			While not RS_Cred2012_Produtos_V2.EOF
				
				Produtos 	= Produtos & "{ "
				Produtos 	= Produtos & """Produto"" : """ & Trim(RS_Cred2012_Produtos_V2("Produto")) & """"
				Produtos 	= Produtos & "} "

				RS_Cred2012_Produtos_V2.MoveNext()
				If not RS_Cred2012_Produtos_V2.EOF Then
					Produtos = Produtos & ","
				End If
			Wend
			RS_Cred2012_Produtos_V2.Close

			Produtos 	= Produtos & "],"
		End If

		'=======================================================================
		' Verificando se empresa já preencheu UNIVERSIDADES
		SQL_Empresa_Universidade = 	"Select " &_
									"	ID_Relacionamento_Cadastro " &_
									"From Relacionamento_Cadastro " &_
									"where  " &_
									"	ID_Empresa = " & ID_Empresa & " " &_
									"	AND ID_Tipo_Credenciamento = 13 " &_
									"	AND ID_Edicao = " & Session("cliente_edicao")

		Set RS_Empresa_Universidade = Server.CreateObject("ADODB.Recordset")
		RS_Empresa_Universidade.Open SQL_Empresa_Universidade, Conexao
		
		empresa_em_universidade = "false"
		If not RS_Empresa_Universidade.BOF or not RS_Empresa_Universidade.EOF Then
			empresa_em_universidade = "true"
			RS_Empresa_Universidade.Close
		End If
		
		'=======================================================================
		' Retornando JSON
		'=======================================================================
		
		%>
		{ 
        "ID_Empresa" : "<%=ID_Empresa%>",
        "ID_Formulario" : "<%=ID_Formulario%>",
		"Razao" : "<%=Razao%>",
		"Fantasia" : "<%=Fantasia%>",
		"Sigla" : "<%=Sigla%>",
		"Produto" : "<%=Produto%>",
		"CEP" : "<%=CEP%>",
		"Endereco" : "<%=Endereco%>",
		"Numero" : "<%=Numero%>",
        "Complemento" : "<%=Complemento%>",
        "Bairro" : "<%=Bairro%>",
		"Cidade" : "<%=Cidade%>",
		"UF" : "<%=UF%>",
		"Pais" : "<%=Pais%>",
		"Site" : "<%=Site%>",
        "CPF" : "<%=CPF%>",
		"NomeF" : "<%=NomeF%>",
		"NomeCredencialF" : "<%=NomeCredF%>",
		"DTNasc" : "<%=DtNasc%>",
		"Cargo" : "<%=CargoF%>",
		"Departamento" : "<%=DeptoF%>",
		<%=Telefones%>
		<%=Ramos%>
		<%=Produtos%>
		"Email" : "",
		"Banco" : "New", 
		"Resultado" : "1",
        "ResultadoTXT" : "Cadastro localizado",
        "Empresa_em_universidade" : "<%=empresa_em_universidade%>"
		}
		<%
	'Else
%>
{ 
Resultado : '0',
ResultadoTXT : 'Cadastro não localizado em nenhuma base'
}
<%
		
	End If	
	
Else 

'=======================================================================
' Cadastro NÃO localizado 
'=======================================================================
%>
{ 
Resultado : '0',
ResultadoTXT : 'Cadastro não localizado em nenhuma base'
}
<%

End If

Conexao.Close
%>