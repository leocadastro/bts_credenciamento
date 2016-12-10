<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/admin/inc/ado_vbs.inc"-->
<%
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")

	edicao = Limpar_Texto(request("id_edicao"))
	idioma = Limpar_Texto(request("id_idioma"))
	tipo = Limpar_Texto(request("id_tipo"))
	
	'response.write edicao

	SQL_Eventos =   "Select " &_
		  " Ee.ID_Evento, " &_
		  " Ee.ID_Edicao, " &_
		  " E.Nome_PTB as Evento, " &_
		  " Ee.Ano " &_
		  "From Eventos_Edicoes as Ee " &_
		  "Inner Join Eventos as E ON E.ID_Evento = Ee.ID_Evento " &_
		  "Where E.ID_Evento = '" & edicao & "'"
		  
		  'response.write SQL_Eventos & "<hr>"
		  
	Set RS_Eventos = Server.CreateObject("ADODB.Recordset")
	RS_Eventos.Open SQL_Eventos, Conexao
  
  If Not RS_Eventos.Eof Then
  	Nome_Feira = RS_Eventos("Evento")
  End If

	nome_relatorio = replace(Nome_Feira," ","-")

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

arquivo = nome_relatorio
extensao = ".xls"

Filename = arquivo & "_" & data & extensao ' file to read

'response.write Filename
'response.end

Function SQL_Exportar(edicao,idioma,tipo)

	'response.write " .2"
	'response.end

	If idioma = 1 then
		Tipo_Idioma = "PTB"
	ElseIf idioma = 2 then
		Tipo_Idioma = "ESP"
	ElseIf idioma = 3 then
		Tipo_Idioma = "ENG"
	End If

	If tipo <> "" then
		WHERE = " and F.Nome = " & tipo & " "
	End If
	
	If idioma <> "" then
		WHERE = " and Rc.ID_Idioma = " & idioma & " "
	End If
	

	SQL_Exportar = "Select "&_
					"	Distinct"&_
					"	Ev.Nome_" & Tipo_Idioma & " as Evento, "&_
					"	Rc.ID_Idioma, "&_
					"	Rc.Data_Cadastro, "&_
					"	EE.Ano, "&_
					"	F.Nome as Formulario, "&_
					"	V.ID_Visitante,  "&_
					"	V.CPF,  "&_
					"	V.Passaporte,  "&_
					"	V.Nome_Completo,  "&_
					"	V.Nome_Credencial,  "&_
					"	V.Data_Nasc, "&_
					"	V.Sexo, "&_
					"	V.Email as Email1, "&_
					"	V.Newsletter as Newsletter1, "&_
					"	Stuff "&_
					"	( "&_
					"		( "&_
					"		Select "&_
					"			TT.Tipo_" & Tipo_Idioma & " + ' - ', "&_
					"			RT.DDI + ' (', "&_
					"			RT.DDD + ') ', "&_
					"			RT.Numero + '', "&_
					"			RT.Ramal + '; ' "&_
					"		From Relacionamento_Telefones as RT "&_
					"		Inner Join Tipo_Telefone as TT ON TT.ID_Tipo_Telefone = RT.ID_Tipo_Telefone "&_
					"		Where ID_Visitante = V.ID_Visitante "&_
					"		For XML PATH ('') "&_
					"		), 1, 0, '' "&_
					"	) as Telefones_Visitante, "&_
					"	C.Cargo_" & Tipo_Idioma & ", "&_
					"	V.Cargo_Outros, "&_
					"	SC.SubCargo_" & Tipo_Idioma & ", "&_
					"	V.SubCargo_Outros, "&_
					"	D.Depto_" & Tipo_Idioma & ", "&_
					"	V.Depto_Outros, "&_
					"	E.ID_Empresa, "&_
					"	FQ.Funcionarios_Qtde_" & Tipo_Idioma & ", "&_
					"	E.CNPJ, "&_
					"	E.Razao_Social, "&_
					"	E.Nome_Fantasia, "&_
					"	E.Principal_Produto, "&_
					"	E.Site, "&_
					"	E.Email as Email2, "&_
					"	E.Newsletter as Newsletter2, "&_
					"	RE.CEP, "&_
					"	RE.Endereco, "&_
					"	RE.Numero, "&_
					"	RE.Complemento, "&_
					"	RE.Bairro, "&_
					"	RE.Cidade, "&_
					"	UF.Estado, "&_
					"	P.Pais, "&_
					"	E.Presidente, "&_
					"	E.Reitor, "&_
					"	E.Senha, "&_
					"	Stuff "&_
					"	( "&_
					"		( "&_
					"		Select "&_
					"			TT.Tipo_" & Tipo_Idioma & " + ' - ', "&_
					"			RT.DDI + ' (', "&_
					"			RT.DDD + ') ', "&_
					"			RT.Numero + '', "&_
					"			RT.Ramal + '; ' "&_
					"		From Relacionamento_Telefones as RT "&_
					"		Inner Join Tipo_Telefone as TT ON TT.ID_Tipo_Telefone = RT.ID_Tipo_Telefone "&_
					"		Where ID_Empresa = E.ID_Empresa "&_
					"		For XML PATH ('') "&_
					"		), 1, 0, '' "&_
					"	) as Telefones_Empresa, "&_
					"	Stuff "&_
					"	( "&_
					"		( "&_
					"		Select  "&_
					"			Ra.Ramo_" & Tipo_Idioma & "  + '  ', "&_
					"			Rr.Ramo_Outros + '; ' "&_
					"		From Relacionamento_Ramo as Rr "&_
					"		Inner Join RamodeAtividade as Ra ON Ra.ID_Ramo = Rr.ID_Ramo "&_
					"		Where ID_Empresa = E.ID_Empresa "&_
					"		For XML PATH ('') "&_
					"		), 1, 0, '' "&_
					"	) as Ramos, "&_
					"	Stuff "&_
					"	( "&_
					"		( "&_
					"		Select  "&_
					"			Ae.Atividade_" & Tipo_Idioma & " + ' ', "&_
					"			Ra.Atividade_Outros + '; ' "&_
					"		From Relacionamento_Atividade as Ra "&_
					"		Inner Join AtividadeEconomica as Ae ON Ae.ID_Atividade = Ra.ID_Atividade "&_
					"		Where ID_Empresa = E.ID_Empresa "&_
					"		For XML PATH ('') "&_
					"		), 1, 0, '' "&_
					"	) as Atividade "&_
					" From Relacionamento_Cadastro as RC  "&_
					" Inner Join Tipo_Credenciamento as TC ON TC.ID_Tipo_Credenciamento = RC.ID_Tipo_Credenciamento  "&_
					" Inner Join Formularios as F ON F.ID_Formulario = TC.ID_Formulario  "&_
					" Inner Join Eventos_Edicoes as EE ON EE.ID_Edicao = RC.ID_Edicao  "&_
					" Inner Join Eventos as EV ON EV.ID_Evento = EE.ID_Evento  "&_
					" Inner Join Visitantes as V ON V.ID_Visitante = RC.ID_Visitante  "&_
					" Left  Join Empresas as E ON E.ID_Empresa = RC.ID_Empresa  "&_
					" Left  Join Relacionamento_Enderecos as RE ON ((RE.ID_Empresa = RC.ID_Empresa) OR (RE.ID_Visitante = RC.ID_Visitante)) "&_
					" Left  Join UF as UF ON UF.ID_UF = RE.ID_UF "&_
					" Left  Join Pais as P ON P.ID_Pais = RE.ID_Pais "&_
					" Left  Join Cargo as C ON C.ID_Cargo = V.ID_Cargo "&_
					" Left  Join SubCargo as SC ON SC.ID_SubCargo = V.ID_SubCargo "&_
					" Left  Join Depto as D ON D.ID_Depto = V.ID_Depto "&_
					" Left  Join Funcionarios_Qtde as FQ ON FQ.ID_Funcionarios_Qtde = E.ID_Funcionarios_Qtde "&_
					" Left  Join Relacionamento_Ramo as RR ON RR.ID_Empresa = RC.ID_Empresa "&_
					" Where  "&_
					" 	RC.ID_Edicao = " & edicao & "  "&_
					WHERE &_
					" Order By F.Nome"
					
		'response.write("<hr>" & SQL_Exportar & "<hr>")	
		'response.end
End Function

Function Monta_Table(Contador,Ano,Formulario,Idioma,Data_Cadastro,ID_Visitante,CPF,Passaporte,Nome_Completo,Nome_Credencial,Data_Nasc,Sexo,Email1,Newsletter1,Telefones_Visitante,Cargo,Cargo_Outros,SubCargo,SubCargo_Outros,Depto,Depto_Outros,ID_Empresa,Funcionarios_Qtde,CNPJ,Razao_Social,Nome_Fantasia,Principal_Produto,Site,Email2,Newsletter2,CEP,Endereco,Numero,Complemento,Bairro,Cidade,Estado,Pais,Presidente,Reitor,Senha,Telefones_Empresa,Ramos,Atividade)
	'response.write " .3"
	'response.end

	CNPJMask 		= Mid(CNPJ,1,2) & "." & Mid(CNPJ,3,3) & "." & Mid(CNPJ,6,3) & "/" & Mid(CNPJ,9,4) & "-" & Mid(CNPJ,13,2)
	CPFMask			= Mid(CPF,1,3) & "." & Mid(CPF,4,3) & "." & Mid(CPF,7,3) & "-" & Mid(CPF,10,2)
	CEPMask			= Mid(CEP,1,5) & "-" & Mid(CEP,4,3)
	Data_NascMask	= Mid(Data_Nasc,1,2) & "/" & Mid(Data_Nasc,3,2) & "/" & Mid(Data_Nasc,5,4)  

	response.write  "<tr>"&_
					"<td nowrap>" & Contador & "</td>"&_
					"<td nowrap>" & Ano & "</td>"&_
					"<td nowrap>" & Formulario & "</td>"&_
					"<td nowrap>" & Idioma & "</td>"&_
					"<td nowrap>" & Data_Cadastro & "</td>"&_
					"<td nowrap>" & ID_Visitante & "</td>"&_
					"<td nowrap>" & CPFMask & "</td>"&_
					"<td nowrap>" & Passaporte & "</td>"&_
					"<td nowrap>" & Nome_Completo & "</td>"&_
					"<td nowrap>" & Nome_Credencial & "</td>"&_
					"<td nowrap>" & Data_NascMask & "</td>"&_
					"<td nowrap>" & Sexo & "</td>"&_
					"<td nowrap>" & Email1 & "</td>"&_
					"<td nowrap>" & Newsletter1 & "</td>"&_
					"<td nowrap>" & Telefones_Visitante & "</td>"&_
					"<td nowrap>" & Cargo & "</td>"&_
					"<td nowrap>" & Cargo_Outros & "</td>"&_
					"<td nowrap>" & SubCargo & "</td>"&_
					"<td nowrap>" & SubCargo_Outros & "</td>"&_
					"<td nowrap>" & Depto & "</td>"&_
					"<td nowrap>" & Depto_Outros & "</td>"&_
					"<td nowrap>" & ID_Empresa & "</td>"&_
					"<td nowrap>" & Funcionarios_Qtde & "</td>"&_
					"<td nowrap>" & CNPJMask & "</td>"&_
					"<td nowrap>" & Razao_Social & "</td>"&_
					"<td nowrap>" & Nome_Fantasia & "</td>"&_
					"<td nowrap>" & Principal_Produto & "</td>"&_
					"<td nowrap>" & Site & "</td>"&_
					"<td nowrap>" & Email2 & "</td>"&_
					"<td nowrap>" & Newsletter2 & "</td>"&_
					"<td nowrap>" & CEPMask & "</td>"&_
					"<td nowrap>" & Endereco & "</td>"&_
					"<td nowrap>" & Numero & "</td>"&_
					"<td nowrap>" & Complemento & "</td>"&_
					"<td nowrap>" & Bairro & "</td>"&_
					"<td nowrap>" & Cidade & "</td>"&_
					"<td nowrap>" & Estado & "</td>"&_
					"<td nowrap>" & Pais & "</td>"&_
					"<td nowrap>" & Presidente & "</td>"&_
					"<td nowrap>" & Reitor & "</td>"&_
					"<td nowrap>" & Senha & "</td>"&_
					"<td nowrap>" & Telefones_Empresa & "</td>"&_
					"<td nowrap>" & Ramos & "</td>"&_
					"<td nowrap>" & Atividade & "</td>"&_
				"</tr>"
End Function

Function Monta_Excel(edicao,idioma,tipo)
	'response.write " .1"
	'response.end
	Set RS_Exportar = Server.CreateObject("ADODB.Recordset")
	RS_Exportar.Open SQL_Exportar(edicao,idioma,tipo), Conexao, 3
	
	If idioma = 1 then
		Tipo_Idioma = "PTB"
	ElseIf idioma = 2 then
		Tipo_Idioma = "ESP"
	ElseIf idioma = 3 then
		Tipo_Idioma = "ENG"
	End If
	
	If Not RS_Exportar.Eof Then
		While Not RS_Exportar.Eof
		
		Contador			= Contador + 1
		Ano 				= RS_Exportar("Ano")
		Formulario 			= RS_Exportar("Formulario")
		Data_Cadastro		= RS_Exportar("Data_Cadastro")
		ID_Visitante 		= RS_Exportar("ID_Visitante")
		CPF 				= RS_Exportar("CPF")
		Passaporte 			= RS_Exportar("Passaporte")
		Nome_Completo 		= RS_Exportar("Nome_Completo")
		Nome_Credencial 	= RS_Exportar("Nome_Credencial")
		Data_Nasc 			= RS_Exportar("Data_Nasc")
		Sexo 				= RS_Exportar("Sexo")
		Email1 				= RS_Exportar("Email1")
		Newsletter1			= RS_Exportar("Newsletter1")
		Telefones_Visitante = RS_Exportar("Telefones_Visitante")
		Cargo 				= RS_Exportar("Cargo_" & Tipo_Idioma)
		Cargo_Outros 		= RS_Exportar("Cargo_Outros")
		SubCargo 			= RS_Exportar("SubCargo_" & Tipo_Idioma)
		SubCargo_Outros 	= RS_Exportar("SubCargo_Outros")
		Depto 				= RS_Exportar("Depto_" & Tipo_Idioma)
		Depto_Outros 		= RS_Exportar("Depto_Outros")
		ID_Empresa 			= RS_Exportar("ID_Empresa")
		Funcionarios_Qtde 	= RS_Exportar("Funcionarios_Qtde_" & Tipo_Idioma)
		CNPJ 				= RS_Exportar("CNPJ")
		Razao_Social 		= RS_Exportar("Razao_Social")
		Nome_Fantasia 		= RS_Exportar("Nome_Fantasia")
		Principal_Produto	= RS_Exportar("Principal_Produto")
		Site 				= RS_Exportar("Site")
		Email2				= RS_Exportar("Email2")
		Newsletter2 		= RS_Exportar("Newsletter2")
		CEP 				= RS_Exportar("CEP")
		Endereco 			= RS_Exportar("Endereco")
		Numero 				= RS_Exportar("Numero")
		Complemento 		= RS_Exportar("Complemento")
		Bairro 				= RS_Exportar("Bairro")
		Cidade				= RS_Exportar("Cidade")
		Estado				= RS_Exportar("Estado")
		Pais 				= RS_Exportar("Pais")
		Presidente 			= RS_Exportar("Presidente")
		Reitor 				= RS_Exportar("Reitor")
		Senha 				= RS_Exportar("Senha")
		Telefones_Empresa 	= RS_Exportar("Telefones_Empresa")
		Ramos 				= RS_Exportar("Ramos")
		Atividade 			= RS_Exportar("Atividade")
		
			'response.write " .4 <hr>"
			Monta_Table Evento,Ano,Formulario,Tipo_Idioma,Data_Cadastro,ID_Visitante,CPF,Passaporte,Nome_Completo,Nome_Credencial,Data_Nasc,Sexo,Email1,Newsletter1,Telefones_Visitante,Cargo,Cargo_Outros,SubCargo,SubCargo_Outros,Depto,Depto_Outros,ID_Empresa,Funcionarios_Qtde,CNPJ,Razao_Social,Nome_Fantasia,Principal_Produto,Site,Email2,Newsletter2,CEP,Endereco,Numero,Complemento,Bairro,Cidade,Estado,Pais,Presidente,Reitor,Senha,Telefones_Empresa,Ramos,Atividade
			'response.write Excel
			'response.end
			
		RS_Exportar.MoveNext
		Wend
		RS_Exportar.Close
	End IF
End Function

	Response.ContentType="application/vnd.ms-excel"
	response.AddHeader "content-disposition", "attachment; filename=" & Filename
%>
<table>
	<tr>
        <td>N</td>
        <td>Ano</td>
        <td>Formulario</td>
        <td>Idioma</td>
        <td>Data_Cadastro</td>
        <td>ID_Visitante</td>
        <td>CPF</td>
        <td>Passaporte</td>
        <td>Nome_Completo</td>
        <td>Nome_Credencial</td>
        <td>Data_Nasc</td>
        <td>Sexo</td>
        <td>Email</td>
        <td>Newsletter</td>
        <td>Telefones_Visitante</td>
        <td>Cargo</td>
        <td>Cargo_Outros</td>
        <td>SubCargo</td>
        <td>SubCargo_Outros</td>
        <td>Depto</td>
        <td>Depto_Outros</td>
        <td>ID_Empresa</td>
        <td>Funcionarios_Qtde</td>
        <td>CNPJ</td>
        <td>Razao_Social</td>
        <td>Nome_Fantasia</td>
        <td>Principal_Produto</td>
        <td>Site</td>
        <td>Email</td>
        <td>Newsletter</td>
        <td>CEP</td>
        <td>Endereco</td>
        <td>Numero</td>
        <td>Complemento</td>
        <td>Bairro</td>
        <td>Cidade</td>
        <td>Estado</td>
        <td>Pais</td>
        <td>Presidente</td>
        <td>Reitor</td>
        <td>Senha</td>
        <td>Telefones_Empresa</td>
        <td>Ramos</td>
        <td>Atividade</td>
	</tr>
	<%
	'response.write idioma
	'response.end
	
	If idioma = "" then
	'response.write " 1"
	'response.end
		response.write Monta_Excel(edicao,1,tipo)
		response.write Monta_Excel(edicao,2,tipo)
		response.write Monta_Excel(edicao,3,tipo)
	Else
	'response.write " 2"
	'response.end
		response.write Monta_Excel(edicao,idioma,tipo)
	End If
	%>
</table>