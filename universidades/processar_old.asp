<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%> 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Credenciamento <%=Year(Now())%> - BTS Informa</title>
<link href="/css/estilos.css" rel="stylesheet" type="text/css">
</head>
<body style="color:#666">
<%
response.Buffer = True
response.Expires = -1
response.AddHeader "Cache-Control", "no-cache"
response.AddHeader "Pragma", "no-cache" 
%>
<!--#include virtual="/admin/inc/gravar_limpar_texto.asp"-->
<!--#include virtual="/admin/inc/acentos_htm.asp"-->
<!--#include virtual="/admin/inc/texto_caixaAltaBaixa.asp"-->
<!--#include virtual="/scripts/enviar_email.asp"-->
<%
response.Charset = "utf-8" 
response.ContentType = "text/html" 

'=======================================================================
'	Select 
'		ID_Formulario
'		,Nome
'	FROM Formularios
' **** Resultado
'	ID_Formulario - Nome
'	1 - Empresa
'	2 - Entidades
'	3 - Imprensa
'	4 - Pessoa Física
'	5 - Universidades
'	6 - Alunos

ID_Formulario	=	5 ' Universidades
'=======================================================================

'For Each item In Request.Form
'	Response.Write "" & item & " - Value: " & Request.Form(item) & "<BR />"
'Next
	
Erro = ""

%>
{
    "Valida": {
	    "Erro": [
<%
' Limpando campos Hidden
id_edicao		= limpar_texto(Request("id_edicao"))
id_idioma		= limpar_texto(Request("id_idioma"))
id_tipo			= limpar_texto(Request("id_tipo"))
origem_cnpj		= limpar_texto(Request("origem_cnpj"))
origem_cpf		= limpar_texto(Request("origem_cpf"))
id_empresa		= limpar_texto(Request("id_empresa"))
id_visitante	= limpar_texto(Request("id_visitante"))

' Limpando o CNPJ
CNPJ 			= limpar_texto(Request("frmCNPJ"))
CNPJ 			= Replace(CNPJ,".","")
CNPJ 			= Replace(CNPJ,"-","")
CNPJ 			= Replace(CNPJ,"/","")
If Len(CNPJ) <> 14 AND Session("cliente_idioma") = 1 Then
	Erro = Erro & "{ ""Erro"", ""01 - CNPJ Incorreto"" },"
End If

Razao			= limpar_texto(Request("frmRazao"))			' Value: BTS INFORMA
Fantasia		= limpar_texto(Request("frmFantasia"))		' Value: BTS INFORMA
Resp 			= limpar_texto(Request("frmResp"))			' Value: BTS INFORMA
Site 			= limpar_texto(Request("frmSite"))			' Value: BTS INFORMA
Senha			= limpar_texto(Request("frmSenha"))			' Value: BTS INFORMA

' Limpando o CEP
CEP				= limpar_texto(Request("frmCEP"))			' Value: 04583-110
CEP				= Replace(CEP,"-","")						' Value: 04583-110
If Len(CEP) <> 8 AND Session("cliente_idioma") = 1 Then
	Erro = Erro & "{ ""Erro"", ""03 - CEP Incorreto"" },"
End If

Endereco	 	= limpar_texto(Request("frmEndereco"))		' Value: AV. DR. CHUCRI ZAIDAN
Numero	 		= limpar_texto(Request("frmNumero"))		' Value: 80
Complemento		= limpar_texto(Request("frmComplemento")) 	' Value: 2º ANDAR
Bairro			= limpar_texto(Request("frmBairro")) 		' Value: MORUMBI
Cidade 			= limpar_texto(Request("frmCidade"))		' Value: SÃO PAULO
Estado 			= limpar_texto(Request("frmEstado"))		' Value: 26
Pais			= limpar_texto(Request("frmPais")) 			' Value: 23

Interesse		= limpar_texto(Request("frmInteresse")) 	' Value: 1

' Limpando o CPF
CPF 			= limpar_texto(Request("frmCPF"))
CPF				= Replace(CPF,".","")
CPF				= Replace(CPF,"-","")
If Len(CPF) <> 11 AND Session("cliente_idioma") = 1 Then
	Erro = Erro & "{ ""Erro"", ""02 - CPF Incorreto"" },"
End If

Passaporte		= limpar_texto(Request("frmPassaporte"))	' Numero do Passaporte
Nome 			= limpar_texto(Request("frmNome"))			' Value: MONICA
NmCracha 		= Left(limpar_texto(Request("frmNmCracha")),27) ' Value: MONICA

' Limpando o Data de Nascimento
DtNasc			= limpar_texto(Request("frmDtNasc")) 		' Value: 28/01/1941
DtNasc			= Replace(DtNasc,"/","") 					' Value: 28/01/1941
If Len(DtNasc) <> 8 Then
	 Erro = Erro & "{ ""Erro"", ""04 - Data de Nascimento Incorreta"" },"
End If

Sexo			= limpar_texto(Request("frmSexo"))			' Value: 1
ID_Cargo		= limpar_texto(Request("frmCargo")) 		' Value: 11
CargoOutros		= limpar_texto(Request("frmCargoOutros"))	' Value: ''
'If CargoOutros = "-1" Then CargoOutros = "NULL"

ID_Depto		= limpar_texto(Request("frmDepto"))			' Value: 1
DeptoOutros		= limpar_texto(Request("frmDeptoOutros"))	' Value: ''
SubCargo 		= limpar_texto(Request("frmSubCargo"))		' Value: 1
If SubCargo = "-" Then SubCargo = "NULL"

DDI				= limpar_texto(Request("frmDDI")) 			' Value: 05
DDD				= limpar_texto(Request("frmDDD")) 			' Value: 222

' Limpando o Telefone 1
Telefone		= limpar_texto(Request("frmTelefone")) 		' Value: 2875-8752
Telefone		= Replace(Telefone,"-","")			 		' Value: 2875-8752
If Len(Telefone) <> 8 AND Session("cliente_idioma") = 1 Then
	 Erro = Erro & "{ ""Erro"", ""05 - Telefone Incorreto"" },"
End If
TelefoneTipo 	= limpar_texto(Request("frmTipo"))			' Value: 1

' Tratar SMS
TelefoneSMS		= limpar_texto(Request("frmSMS"))			' Value: 1
If Len(Trim(TelefoneSMS)) = 0 Then TelefoneSMS = 0

Ramal			= limpar_texto(Request("frmRamal"))			' Value: 12345

DDI2 			= limpar_texto(Request("frmDDI2"))			' Value: 11
DDD2			= limpar_texto(Request("frmDDD2")) 			' Value: 111

' Limpando o Telefone 2
Telefone2		= limpar_texto(Request("frmTelefone2"))		' Value: 1111-1111
Telefone2		= Replace(Telefone2,"-","")					' Value: 1111-1111
If Len(Telefone2) <> 8 AND Session("cliente_idioma") = 1 Then
	 Erro = Erro & "{ ""Erro"", ""06 - Telefone 2 Incorreto"" },"
End If
TelefoneTipo2	= limpar_texto(Request("frmTipo2"))			' Value: 2

' Tratar SMS2
TelefoneSMS2	= limpar_texto(Request("frmSMS2"))			' Value: 1
If Len(Trim(TelefoneSMS2)) = 0 Then TelefoneSMS2 = 0

Ramal2			= limpar_texto(Request("frmRamal2"))		' Value: 12345


DDIEmpresa		= limpar_texto(Request("frmDDIEmpresa"))			' Value: 11
DDDEmpresa		= limpar_texto(Request("frmDDDEmpresa")) 			' Value: 111

' Limpando o Telefone 2
TelefoneEmpresa	= limpar_texto(Request("frmTelefoneEmpresa"))		' Value: 1111-1111
TelefoneEmpresa	= Replace(TelefoneEmpresa,"-","")					' Value: 1111-1111
If Len(TelefoneEmpresa) <> 8 AND id_idioma = 1 Then
	 Erro = Erro & "{ ""Erro"", ""06 - Telefone Empresa Incorreto"" },"
End If
TelefoneTipoEmpresa	= limpar_texto(Request("frmTipoEmpresa"))' Value: 2

' Tratar SMS2
TelefoneSMSEmpresa	= limpar_texto(Request("frmSMSEmpresa"))' Value: 1
If Len(Trim(TelefoneSMSEmpresa)) = 0 Then TelefoneSMSEmpresa = 0

RamalEmpresa			= limpar_texto(Request("frmRamalEmpresa"))		' Value: 12345

Email 			= Lcase(limpar_texto(Request("frmEmail")))	' Value: monica.mendes@btsmedia.biz
EmailConf 		= Lcase(limpar_texto(Request("frmEmailConf"))) ' Value: monica.mendes@btsmedia.biz
If Email <> EmailConf Then
	 Erro = Erro & "{ ""Erro"", ""07 - E-mails não são iguais"" },"
End If

Newsletter		= limpar_texto(Request("frmNewsletter"))
If Len(Trim(Newsletter)) = 0 Then
	Newsletter = 0
End If	

TotPerguntas 	= limpar_texto(Request("frmTotPerguntas"))

' Retorno das Validações do form
If Len(Erro) > 0 Then
	Retorno = "0"
Else
	Retorno = "1"
End If
%>
	<%=Erro%>
    ]
},
	"Retorno" : "<%=Retorno%>" 
}
<%

'=======================================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'=======================================================================

'****************************
' Descricao dos PASSOS
' 1) Utilizar Campos HIDDEN
' 2) Verificar se a Empresa EXISTE
'	A) Caso SIM - Atualizar
'	B) Senão 	- Cadastrar
' 3) Verificar se o Visitante EXISTE
'	A) Caso SIM	- Atualizar
'	B) Senão	- Cadastrar
' 4) Gravar Relacionamento CADASTRO
' 5) Disparar email de confirmação
' 5) Postar Info´s para 
'****************************


'=======================================================================
	' 2) A - [ATUALIZAR] - Empresa existe no banco Atual
	If origem_cnpj = "novo" OR Len(id_empresa) > 0 Then
		SQL_Verificar_Empresa	=	"Select " &_
									"	ID_Empresa " &_
									"From Empresas " &_
									"Where " &_
									"	ID_Empresa = " & id_empresa 

		Set RS_Verificar_Empresa = Server.CreateObject("ADODB.Recordset")
		RS_Verificar_Empresa.Open SQL_Verificar_Empresa, Conexao

		'Se existe Atualizar
		If not RS_Verificar_Empresa.BOF or not RS_Verificar_Empresa.EOF Then
			SQL_Atualizar_Empresa = "Update Empresas " &_
									"Set " &_
									"	CNPJ 					= Upper(dbo.sp_rm_accent_pt_latin1('" & CNPJ 		& "')), " &_
									"	Razao_Social 			= Upper(dbo.sp_rm_accent_pt_latin1('" & Razao 		& "')), " &_
									"	Fantasia 				= Upper(dbo.sp_rm_accent_pt_latin1('" & Sigla 		& "')), " &_
									"	Site 					= Upper(dbo.sp_rm_accent_pt_latin1('" & Site 		& "')), " &_
									"	Email 					= Upper(dbo.sp_rm_accent_pt_latin1('" & Email 		& "')), " &_
									"	Newsletter 				= "  & Newsletter 	& ", " &_
									"	Data_Atualizacao 		= getDate() "
			Set RS_Atualizar_Empresa = Server.CreateObject("ADODB.Recordset")
			RS_Atualizar_Empresa.Open SQL_Atualizar_Empresa, Conexao
		End If

		' #################################################################################
		' FALTA
		' - receber id_endereco selecionado
		' - atualizar endereço
		' - receber id_telefone
		' - atualizar telefone
		' #################################################################################
	End If

'=======================================================================
	'2) B - [CADASTRAR] - Empresa veio do banco anterior ou não existe
	If (origem_cnpj = "" OR origem_cnpj = "old") AND Len(Trim(id_empresa)) = 0 Then
		' #################################################################################
		' Inserir EMPRESA
		'	Inserir Endereço empresa
		'	Inserir Interesses na feira
		' Inserir VISITANTE
		'	Inserir Telefones 1 e 2
		' Inserir Relacionamento Feira > Empresa > Visitante
		' Inserir Perguntas com o ID do Relacionamento
		' #################################################################################
		
		'=======================================================================
		' Inserir EMPRESA
		SQL_Cad_Empresa = 	"SET NOCOUNT ON;" &_
							" " & vbCrLf & " " &_
							"INSERT INTO Empresas " &_
							"	(ID_Formulario " &_
							"	,CNPJ " &_
							"	,Razao_Social " &_
							"	,Sigla " &_
							"	,Site " &_
							"	,Senha " &_
							"	,Coordenador_Curso) " &_
							"VALUES " &_
							"	(" & ID_Formulario & ", " &_
							"	Upper(dbo.sp_rm_accent_pt_latin1('" & Left(CNPJ, 14) & "')), " &_
							"	Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Razao, 150) & "')), " &_
							"	Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Sigla, 20) & "')), " &_
							"	Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Site, 100) & "')), " &_
							"	Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Senha, 20) & "')), " &_
							"	Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Resp, 100) & "')) " &_
							"	); " &_
							" " & vbCrLf & " " &_
							"SELECT @@Identity as NovoID; "

		 response.write("<hr><b>SQL_Cad_Empresa</b><br>" & SQL_Cad_Empresa & "<hr>")

		'Executando Gravação com Retorno do ID
		Set RS_Cad_Empresa = Conexao.Execute(SQL_Cad_Empresa)
		Novo_ID_Empresa = RS_Cad_Empresa.Fields("NovoID").value
		Set RS_Cad_Empresa = Nothing
		response.write("Novo_ID_Empresa: " & Novo_ID_Empresa)
		'=======================================================================

		'=======================================================================
		' Inserir Relacionamento Cadastro
		SQL_Rel_Cadastro = 	"SET NOCOUNT ON;" &_
							" " & vbCrLf & " " &_
							"INSERT INTO Relacionamento_Cadastro " &_
							"	(ID_Idioma " &_
							"	,ID_Edicao " &_
							"	,ID_Tipo_Credenciamento " &_
							"	,ID_Empresa) " &_
							"VALUES " &_
							"	(" & id_idioma & ", " &_
							"	 " & id_edicao & ", " &_
							"	 " & id_tipo & ", " &_
							"	 " & Novo_ID_Empresa & ");" &_
							" " & vbCrLf & " " &_
							"SELECT @@Identity as NovoID; "

		 response.write("<hr><b>SQL_Rel_Cadastro</b><br>" & SQL_Rel_Cadastro & "<hr>")

		' Executando Gravação com Retorno do ID
		Set RS_Rel_Cadastro = Conexao.Execute(SQL_Rel_Cadastro)
		Novo_ID_Rel_Cadastro = RS_Rel_Cadastro.Fields("NovoID").value
		Set RS_Rel_Cadastro = Nothing
		response.write("Novo_ID_Rel_Cadastro: " & Novo_ID_Rel_Cadastro)
		'=======================================================================

		'=======================================================================
		' Inserir Ramos Selecionados
		Lista_Ramos = Split(OptRamo,",")
		For i = Lbound(Lista_Ramos) to Ubound(Lista_Ramos)
			response.write("i: " & i & " - v: " & Lista_Ramos(i) & "<br>")
			
			If Trim(Lista_Ramos(i)) <> "-1" Then

				SQL_Cad_Ramo = 	"INSERT INTO Relacionamento_Ramo " &_
								"	(ID_Empresa " &_
								"	,ID_Ramo) " &_
								"VALUES " &_
								"	(" & Novo_ID_Empresa & ", " &_
								"	" & Lista_Ramos(i) & "); "

				 response.write("<hr><b>SQL_Cad_Ramo</b><br>" & SQL_Cad_Ramo & "<hr>")
				' Executando Gravação
				Set RS_Cad_Ramo = Conexao.Execute(SQL_Cad_Ramo)
			
			End If

		Next
		response.write("OptRamoOutros: " & OptRamoOutros)
		' Se não vier vazio gravar
		If Len(Trim(OptRamoOutros)) > 0 Then
		SQL_Cad_Ramo_Outros = 	"INSERT INTO Relacionamento_Ramo " &_
								"	(ID_Empresa " &_
								"	,Ramo_Outros " &_
								"	,ID_Ramo) " &_
								"VALUES " &_
								"	(" & Novo_ID_Empresa & ", " &_
								"	'" & Left(OptRamoOutros,255) & "', " &_ 
								"	-1); "

		 response.write("<hr><b>SQL_Cad_Ramo_Outros</b><br>" & SQL_Cad_Ramo_Outros & "<hr>")
		' Executando Gravação
		Set RS_Cad_Ramo_Outros = Conexao.Execute(SQL_Cad_Ramo_Outros)
		End If
		'=======================================================================

		'=======================================================================
		' Inserir Os Interesses Selecionados
		response.write("Lista_Interesses: " & Interesse)
		Lista_Interesses = Split(Interesse,",")
		For i = Lbound(Lista_Interesses) to Ubound(Lista_Interesses)
			response.write("i: " & i & " - v: " & Lista_Interesses(i) & "<br>")
			SQL_Cad_Interesse = 	"INSERT INTO Relacionamento_InteresseFeira " &_
									"	(ID_Relacionamento_Cadastro " &_
									"	,ID_InteresseFeira) " &_
									"VALUES " &_
									"	(" & Novo_ID_Rel_Cadastro & ", " &_
									"	" & Lista_Interesses(i) & "); "

			 response.write("<hr><b>SQL_Cad_Interesse</b><br>" & SQL_Cad_Interesse & "<hr>")
			' Executando Gravação
			Set RS_Cad_Interesse = Conexao.Execute(SQL_Cad_Interesse)
		Next
		'=======================================================================

		'=======================================================================
		' Inserir Endereco da EMPRESA
		SQL_Cad_End_Empresa = 	"INSERT INTO Relacionamento_Enderecos " &_
								"	( " &_
								"	ID_Empresa " &_
								"	,CEP " &_
								"	,Endereco " &_
								"	,Numero " &_
								"	,Complemento " &_
								"	,Bairro " &_
								"	,Cidade " &_
								"	,ID_UF " &_
								"	,ID_Pais " &_
								"	) " &_
								"VALUES " &_
								"	( " &_
								"	" & Novo_ID_Empresa & " " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(CEP,12) & "')) " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Endereco,200) & "')) " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Numero,20) & "')) " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Complemento,50) & "')) " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Bairro, 200) & "')) " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Cidade, 200) & "')) " &_
								"	," & Estado & " " &_
								"	," & Pais & " " &_
								"	);"
		 response.write("<hr><b>SQL_Cad_End_Empresa</b><br>" & SQL_Cad_End_Empresa & "<hr>")
		' Executando Gravação
		Set RS_Cad_End_Empresa = Conexao.Execute(SQL_Cad_End_Empresa)
		'=======================================================================


		'=======================================================================
		' Inserir TELEFONES DO EMPRESA
		SQL_Cad_Tel_Empresa = 	"INSERT INTO Relacionamento_Telefones " &_
									"	( " &_
									"	ID_Empresa " &_
									"	,ID_Tipo_Telefone " &_
									"	,DDI " &_
									"	,DDD " &_
									"	,Numero " &_
									"	,Ramal " &_
									"	,SMS " &_
									"	) " &_
									"VALUES " &_
									"	( " &_
									"	" & Novo_ID_Empresa & " " &_
									"	," & TelefoneTipoEmpresa & " " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(DDIEmpresa,5) & "')) " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(DDDEmpresa,5) & "')) " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(TelefoneEmpresa,15) & "')) " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(RamalEmpresa,5) & "')) " &_
									"	," & TelefoneSMSEmpresa & " " &_
									"	)"
		 response.write("<hr><b>SQL_Cad_Tel_Empresa</b><br>" & SQL_Cad_Tel_Empresa & "<hr>")
		' Executando Gravação
		Set RS_Cad_Tel_Empresa = Conexao.Execute(SQL_Cad_Tel_Empresa)
		'=======================================================================
		

		'=======================================================================
		' Inserir VISITANTE
		SQL_Cad_Visitante = 	"SET NOCOUNT ON;" &_
								" " & vbCrLf & " " &_
								"INSERT INTO Visitantes " &_
								"	( " &_
								"	CPF " &_
								"	,Passaporte " &_
								"	,Nome_Completo " &_
								"	,Nome_Credencial " &_
								"	,Data_Nasc " &_
								"	,Sexo " &_
								"	,Email " &_
								"	,Newsletter " &_
								"	,ID_Cargo " &_
								"	,Cargo_Outros " &_
								"	,ID_SubCargo " &_
								"	,SubCargo_Outros " &_
								"	,ID_Depto " &_
								"	,Depto_Outros " &_
								"	) " &_
								"VALUES " &_
								"	( " &_
								"	'" & Left(CPF,11) & "' " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Passaporte,50) &"')) " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Nome,150) &"')) " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(NmCracha,27) & "')) " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(DtNasc,8) & "')) " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Sexo,1) & "')) " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Email,150) & "')) " &_
								"	," & Newsletter & " " &_
								"	," & Cargo & " " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(CargoOutros,50) & "')) " &_
								"	," & SubCargo & " " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(SubCargoOutros,50) & "')) " &_
								"	," & Depto & " " &_
								"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(DeptoOutros,50) & "')) " &_
								"	); " &_
								" " & vbCrLf & " " &_
								"SELECT @@Identity as NovoID; "

		response.write("<hr><b>SQL_Cad_Visitante</b><br>" & SQL_Cad_Visitante & "<hr>")
		' Executando Gravação com Retorno do ID
		Set RS_Cad_Visitante = Conexao.Execute(SQL_Cad_Visitante)
		Novo_ID_Visitante = RS_Cad_Visitante.Fields("NovoID").value
		Set RS_Cad_Visitante = Nothing
		response.write("Novo_ID_Visitante: " & Novo_ID_Visitante)
		'=======================================================================

		'=======================================================================
		' Inserir TELEFONES DO VISITANTE
		SQL_Cad_Tel_Visitante = 	"INSERT INTO Relacionamento_Telefones " &_
									"	( " &_
									"	ID_Visitante " &_
									"	,ID_Tipo_Telefone " &_
									"	,DDI " &_
									"	,DDD " &_
									"	,Numero " &_
									"	,Ramal " &_
									"	,SMS " &_
									"	) " &_
									"VALUES " &_
									"	( " &_
									"	" & Novo_ID_Visitante & " " &_
									"	," & TelefoneTipo & " " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(DDI,3) & "')) " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(DDD,3) & "')) " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Telefone,15) & "')) " &_
									"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Ramal,5) & "')) " &_
									"	," & TelefoneSMS & " " &_
									"	)"
		response.write("<hr><b>SQL_Cad_Tel_Visitante</b><hr>" & SQL_Cad_Tel_Visitante & "<hr>")
		' Executando Gravação
		Set RS_Cad_Tel_Visitante = Conexao.Execute(SQL_Cad_Tel_Visitante)
		'=======================================================================

		If Len(Telefone2) > 0 Then

			'=======================================================================
			' Inserir TELEFONES DO VISITANTE
			SQL_Cad_Tel_Visitante = 	"INSERT INTO Relacionamento_Telefones " &_
										"	( " &_
										"	ID_Visitante " &_
										"	,ID_Tipo_Telefone " &_
										"	,DDI " &_
										"	,DDD " &_
										"	,Numero " &_
										"	,Ramal " &_
										"	,SMS " &_
										"	) " &_
										"VALUES " &_
										"	( " &_
										"	" & Novo_ID_Visitante & " " &_
										"	," & TelefoneTipo2 & " " &_
										"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(DDI2,3) & "')) " &_
										"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(DDD2,3) & "')) " &_
										"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Telefone2,15) & "')) " &_
										"	,Upper(dbo.sp_rm_accent_pt_latin1('" & Left(Ramal2,5) & "')) " &_
										"	," & TelefoneSMS2 & " " &_
										"	)"
			' response.write("<hr><b>SQL_Cad_Tel_Visitante</b><br>" & SQL_Cad_Tel_Visitante & "<hr>")
			' Executando Gravação
			Set RS_Cad_Tel_Visitante = Conexao.Execute(SQL_Cad_Tel_Visitante)
			'=======================================================================

		End if

		'=======================================================================
		' Inserir PERGUNTAS
		' Existe total de perguntas ?
		If Len(TotPerguntas) > 0 Then
			' Loop na quantidade	
			For x = 1 To TotPerguntas
				ID_Pergunta = limpar_texto(Request("ID_Pergunta_" & x))
				Lista_Perguntas = Split(limpar_texto(Request("frmPergunta_" & x)),",")
				' Loop nos valores
				For y = Lbound(Lista_Perguntas) to Ubound(Lista_Perguntas)
					SQL_Cad_Perguntas = 	"INSERT INTO Relacionamento_Perguntas " &_
											"	( " &_
											"	ID_Relacionamento_Cadastro " &_
											"	,ID_Perguntas " &_
											"	,ID_Opcoes " &_
											"	,Texto " &_
											"	) " &_
											"VALUES " &_
											"	(" &_
											"	" & Novo_ID_Rel_Cadastro & ", " &_
											"	" & ID_Pergunta & ", " &_
											"	" & Lista_Perguntas(y) & ", " &_
											"	'');"
					response.write("<hr><b>SQL_Cad_Perguntas x(" & x & ") / y(" & y & ") / val(" & Lista_Perguntas(y) & ") </b><br>" & SQL_Cad_Perguntas & "<hr>")
					' Executando Gravação
					Set RS_Cad_Perguntas = Conexao.Execute(SQL_Cad_Perguntas)
				Next
			Next
		End If
		'=======================================================================

		'=======================================================================
		' Atualizar RELACIONAMENTO CADASTRO
		SQL_Upd_Rel_Cadastro =	"Update Relacionamento_Cadastro " &_
								"Set " &_
								" 	ID_Visitante = " & Novo_ID_Visitante & " " &_
								"Where " &_
								"	ID_Relacionamento_Cadastro = " & Novo_ID_Rel_Cadastro
		 response.write("<hr><b>SQL_Upd_Rel_Cadastro</b><br>" & SQL_Upd_Rel_Cadastro & "<hr>")
		' Executando Gravação
		Set RS_Upd_Rel_Cadastro = Conexao.Execute(SQL_Upd_Rel_Cadastro)
		'=======================================================================
	End If

	' Enviar EMAIL
	Enviar_Email id_edicao, id_idioma, ID_Formulario, Email, Novo_ID_Rel_Cadastro, CPF, Nome, ID_Cargo, ID_Depto, CNPJ, Razao

Conexao.Close

Response.Clear()

Session("cliente_empresa") 	= Novo_ID_Empresa
Session("cliente_logado") 	= True				' Para a tela de Alunos não redirecionar ao LOGIN
%>
<form id="confirmacao" name="confirmacao" method="POST" action="/alunos/cadastrar.asp">
	<input type="hidden" name="id_edicao" value="<%=id_edicao%>">
	<input type="hidden" name="id_idioma" value="<%=id_idioma%>">
	<input type="hidden" name="id_tipo" value="<%=id_tipo%>">
	<input type="hidden" name="frmID_Cadastro" value="<%=Novo_ID_Rel_Cadastro%>">
	<input type="hidden" name="frmID_Empresa" value="<%=Novo_ID_Empresa%>">
	<input type="hidden" name="frmNome" value="<%=Nome%>">
	<input type="hidden" name="frmCPF" value="<%=CPF%>">
	<input type="hidden" name="frmCargo" value="<%=ID_Cargo%>">
	<input type="hidden" name="frmDepartamento" value="<%=ID_Depto%>">
	<input type="hidden" name="frmCNPJ" value="<%=CNPJ%>">
	<input type="hidden" name="frmRazaoSocial" value="<%=Razao%>">
</form>
<script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
<script language="Javascript">
	$(document).ready(function(){
		$("#confirmacao").submit();
	});
</script>
</body>
</html>