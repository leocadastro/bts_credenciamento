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
	'response.Write("<strong>SQL_Verificar</strong><hr>" & SQL_Verificar & "<hr>")
	Set RS_Cred2012_Empresa = Server.CreateObject("ADODB.Recordset")
	RS_Cred2012_Empresa.CursorType = 0
	RS_Cred2012_Empresa.LockType = 1
	RS_Cred2012_Empresa.Open SQL_Cred2012_Empresa, Conexao

	'=======================================================================
	' Verificando o Retorno da Query SQL_Verificar - Credenciamento_2012
	'=======================================================================
	If RS_Cred2012_Empresa.BOF or RS_Cred2012_Empresa.EOF Then
		'=======================================================================
		' Nao existe - Retorno Empresa nao localizada
		'=======================================================================
	%>
		{ 
			Resultado : '0',
			ResultadoTXT : 'Cadastro não localizado em nenhuma base'
		}
	<%
			
	Else

		'=======================================================================
		' Cadastro localizado no Banco Novo Credenciamento_2012
		' TRATANDO AS VARIAVEIS
		'=======================================================================
		
		' Verificando Produtos cadastrados para uma Edição específica
		SQL_Verificar_Produtos = 	"SELECT " &_ 
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
		Set RS_Cred2012_Verificar_Produtos = Server.CreateObject("ADODB.Recordset")
		RS_Cred2012_Verificar_Produtos.CursorType = 0
		RS_Cred2012_Verificar_Produtos.LockType = 1
		RS_Cred2012_Verificar_Produtos.Open SQL_Verificar_Produtos, Conexao	

		If RS_Cred2012_Verificar_Produtos.BOF or RS_Cred2012_Verificar_Produtos.EOF Then
			'=======================================================================
			' Nao Existe - Retorno Produtos localizados
			'=======================================================================
		%>
			{ 
				Resultado : '0',
				ResultadoTXT : 'Produtos não localizado em nenhuma base'
			}
		<%
		Else
			'=======================================================================
			' Existe - Retorno Produtos localizados
			'=======================================================================

			%>
				{ 
        			"Produtos" : "<%=ID_Empresa%>",
			<%

			While not RS_Cred2012_Verificar_Produtos.EOF
				
				Telefones = Telefones & """DDI" & t & """ : """ & Trim(RS_Cred2012_Empresa_Telefone("DDI")) & ""","
				Telefones = Telefones & """DDD" & t & """ : """ & Trim(RS_Cred2012_Empresa_Telefone("DDD")) & ""","
				Telefones = Telefones & """Fone" & t & """ : """ & Trim(RS_Cred2012_Empresa_Telefone("Numero")) & ""","

				t = t + 1

				RS_Cred2012_Verificar_Produtos.MoveNext()
			Wend
			RS_Cred2012_Verificar_Produtos.Close	
		
		'=======================================================================
		' Retornando JSON
		'=======================================================================
		
		%>
		
		<%

	End If	
	
Else 

'=======================================================================
' Cadastro NÃO localizado 
'=======================================================================
%>

<%

End If

Conexao.Close
%>