<!--#include virtual="/admin/inc/limpar_texto.asp"-->


<!DOCTYPE html PUBLIC
  "-//W3C//DTD XHTML 1.0 Strict//EN"
  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html>
  <head>
    <title>Form Iframe Demo</title>

  </head>
  <body id="iframe-body">


<%

	If Not Session("admin_logado") Then
		Response.Redirect("/admin/")
	End If

	Response.Expires = -1
	Response.Buffer = True

	'// Define o tipo do arquivo que ser? exportado
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=" & NomeArquivo()

	'// Declara as vari?veis
	Dim objConexao, objRetorno
	Dim ID_Pedido_Status
	Dim Ano_Pedido
	Dim Cod_Campo
	Dim Campo
	Dim Campo_Valor
	Dim strSQL
	Dim ID_Evento

	'// Define os valores das variaveis
	ID_Pedido_Status = Limpar_Texto(Request("id_status"))
	Ano_Pedido = Limpar_Texto(Request("ano_pedido"))
	Campo = Limpar_Texto(Request("campo_busca"))
	Campo_Valor = Limpar_Texto(Request("pedido"))
	ID_Evento = Limpar_Texto(Request("ID_Evento"))


	if ID_Pedido_Status = "" then
		status = "in (1,2,3,4)"
	else
		status = " = " & ID_Pedido_Status
	End if

	'// Monta a string SQL
	strSQL = "SELECT pe.Numero_Pedido, vc.ID_Visitante, v2.CPF AS 'CPF_Comprador', vc.Nome_Completo, vc.Nome_Credencial, vc.CPF, vc.Email, pe.Valor_Pedido, pe.Data_Pedido, "
	strSQL = strSQL + "ps.Status_PTB as Status_PTB, COUNT(pc.ID_Carrinho) AS Ingressos, ph.Codigo_Paypal, EL.Nome as NomeLote "
	strSQL = strSQL + "from pedidos pe "
	strSQL = strSQL + "left join Pedidos_Historico ph on pe.Numero_Pedido = ph.Numero_Pedido  and (ph.codigo_autorizacao = 'SUCCESS' or ph.codigo_autorizacao = 'SUCCESSWITHWARNING') "
	strSQL = strSQL + "inner join Pedidos_Carrinho pc on pc.ID_Pedido = pe.ID_Pedido "
	strSQL = strSQL + "inner join Visitantes Vc on Vc.ID_Visitante = pc.id_visitante "
	strSQL = strSQL + "inner join Pedidos_Status ps on pe.Status_Pedido = ps.ID_Pedido_Status "
	strSQL = strSQL + "inner join Visitantes v2 on v2.ID_Visitante = pe.ID_Visitante "
	strSQL = strSQL + "inner join Edicoes_Lote EL on EL.ID_Edicao = pe.ID_Edicao And pe.Data_Pedido between EL.Data_Inicio And EL.Data_Fim "
	strSQL = strSQL + "where pe.ID_Edicao = 60 "
	strSQL = strSQL + "	and pe.Numero_Pedido <> '' And EL.Ativo = 1 "
	strSQL = strSQL + "	and pe.Valor_Pedido > 0 And Pe.Status_Pedido " & status & " "
	strSQL = strSQL + "GROUP BY pe.Numero_Pedido, vc.ID_Visitante, vc.Nome_Completo, vc.Nome_Credencial, vc.CPF, vc.Email, pe.Valor_Pedido, pe.Data_Pedido, "
	strSQL = strSQL + "	ph.data_pagamento, ps.Status_PTB, ph.Codigo_Paypal, v2.CPF, EL.Nome "
	strSQL = strSQL + "order by ph.data_pagamento "

	'Response.Write strSQL
	'Response.End

	'// Abre a conex?o com o banco e executa a string SQL
	'Set objConexao = Server.CreateObject("ADODB.Connection")
	'objConexao.Open Application("cnn")


	Set Conexao_cred = Server.CreateObject("ADODB.Connection")
	Conexao_cred.Open Application("cnn")

	'response.write strSQL
	'response.end
	Set objRetorno = Server.CreateObject("ADODB.Recordset")
				objRetorno.Open strSQL, Conexao_cred.ConnectionString, 3, 3


	'Set objRetorno = Server.CreateObject("ADODB.Recordset")
	'objRetorno.CursorLocation = 3
	'objRetorno.CursorType = 0
	'objRetorno.LockType = 1
		'response.write objConexao.ConnectionString
	'response.end
	'Set objRetorno = objConexao.Execute(strSQL)

	'// Verifica se houve retorno

	'response.write strSql
	'response.end

	If Not objRetorno.EOF Then
		'// Monta as colunas de header do documento?
		Response.Write("<table border=1>")
		Response.Write("<tr>")
		Response.Write("<th>Número do Pedido</th>")
		Response.Write("<th>ID Visitante</th>")
		Response.Write("<th>CPF Comprador</th>")
		Response.Write("<th>Nome</th>")
		Response.Write("<th>CPF</th>")
		Response.Write("<th>E-mail</th>")
		Response.Write("<th>Ingressos</th>")
		Response.Write("<th>Valor</th>")
		Response.Write("<th>Data</th>")
		Response.Write("<th>Status</th>")
		Response.Write("<th>Código Paypal</th>")
		Response.Write("<th>Série Ingresso</th>")
		Response.Write("</tr>")

		'// Retorna as linhas do documento
		n_ped = ""
		tot_ped = 0
		i = 0
		While Not objRetorno.EOF

			if n_ped = objRetorno("Numero_Pedido") then
				tot_ped = tot_ped + 1
			else
				n_ped = objRetorno("Numero_Pedido")
				tot_ped = 1
			end if

			Response.Write("<tr>")
			Response.Write("<td>" & objRetorno("Numero_Pedido") & "</td>")
			Response.Write("<td>" & objRetorno("ID_Visitante") & "</td>")
			Response.Write("<td>" & objRetorno("CPF_Comprador") & "</td>")
			Response.Write("<td nowrap>" & objRetorno("Nome_Completo") & "</td>")
			Response.Write("<td style=""mso-number-format:\@"">" & objRetorno("CPF") & "</td>")
			Response.Write("<td nowrap>" & objRetorno("Email") & "</td>")
			Response.Write("<td>" & objRetorno("Ingressos") & "</td>")
			Response.Write("<td style=""mso-number-format:'0\.00'"">" & objRetorno("Valor_Pedido") & "</td>")
			Response.Write("<td>" & objRetorno("Data_Pedido") & "</td>")
			Response.Write("<td nowrap>" & objRetorno("Status_PTB") & "</td>")
			Response.Write("<td nowrap>" & objRetorno("Codigo_Paypal") & "</td>")
			Response.Write("<td>" & objRetorno("NomeLote") & "</td>")
			Response.Write("</tr>")

			objRetorno.MoveNext()

			i = i + 1
		Wend

		Response.Write("</table>")
	End If

	'// Finaliza os objetos

	'objConexao.Close
	'Set objConexao = Nothing
	Set objRetorno = Nothing

	'// Fun??o para gerar o nome o arquivo com a hora atual
	Function NomeArquivo()
		Dim dia, mes, ano, hora, minuto, segundo
		Dim textoHorario

		dia = Right("0" & Day(Now()), 2)
		mes = Right("0" & Month(Now()), 2)
		ano = Year(Now())
		hora = Right("0" & Hour(Now()), 2)
		minuto = Right("0" & Minute(Now()), 2)
		segundo = Right("0" & Second(Now()), 2)

		textoHorario = ano & mes & dia & "_" & hora & minuto & segundo

		NomeArquivo = "Ingressos_ABF_" & textoHorario & ".xls"
		'NomeArquivo = "Ingressos_ABF_novo.xls"

	End Function

	Response.Flush()

%>

  </body>
</html>
