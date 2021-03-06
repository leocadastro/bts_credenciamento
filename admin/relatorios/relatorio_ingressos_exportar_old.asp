<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<%
	If Not Session("admin_logado") Then
		Response.Redirect("/admin/")
	End If
	
	Response.Expires = -1
	Response.Buffer = True
	
	'// Define o tipo do arquivo que ser� exportado
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=" & NomeArquivo()
	
	'// Declara as vari�veis
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
	
	
	'// Monta a string SQL
	strSQL = "SELECT pe.Numero_Pedido, vc.ID_Visitante, vc.Nome_Completo, vc.Nome_Credencial, vc.CPF, vc.Email, pe.Valor_Pedido, pe.Data_Pedido, 'Pedido Concluido' as Status_PTB, COUNT(pc.ID_Carrinho) AS Ingressos"
	strSQL = strSQL & " from  Pedidos_Historico  "
	strSQL = strSQL & " inner join pedidos pe on pe.Numero_Pedido = Pedidos_Historico.Numero_Pedido "
		strSQL = strSQL & " inner join  Pedidos_Carrinho pc on pc.ID_Pedido = pe.ID_Pedido "
		strSQL = strSQL & " inner join Visitantes Vc on Vc.ID_Visitante = pc.id_visitante "
	strSQL = strSQL & " 	where  pe.ID_Edicao = 56 and  Pedidos_Historico.Data_Pagamento like '%2015%' "
	strSQL = strSQL & " and (Pedidos_Historico.codigo_autorizacao like 'SUCCESS' or Pedidos_Historico.codigo_autorizacao like 'SUCCESSWITHWARNING') "
	strSQL = strSQL & " and Pedidos_Historico.data_pagamento > '2015-04-29 16:00:00.000' and Pedidos_Historico.Numero_Pedido <> '' and Vc.email not like '%paypal%' And Pe.Status_Pedido = " & ID_Pedido_Status & " " 

	'If ID_Pedido_Status <> "" And IsNumeric(ID_Pedido_Status) Then
	'	strSQL = strSQL & " 	AND ped.Status_Pedido = " & ID_Pedido_Status
	'End If
	'If Ano_Pedido <> "" And IsNumeric(Ano_Pedido) Then
	'	strSQL = strSQL & " 	AND YEAR(ped.Data_Pedido) = " & Ano_Pedido
	'End If
	'If Campo <> "" And Campo_Valor <> "" Then
	'	Select Case Campo
	'		Case "numeropedido"
		'		strSQL = strSQL & " 	AND ped.Numero_Pedido LIKE '%" & Campo_Valor & "%'"
		'	Case "nomecomprador"
		'		strSQL = strSQL & " 	AND vis.Nome_Completo LIKE '%" & Campo_Valor & "%'"
		'	Case "cpfcomprador"
		'		strSQL = strSQL & " 	AND vis.CPF LIKE '%" & Campo_Valor & "%'"
		'End Select
	'End If
	strSQL = strSQL & " GROUP BY pe.Numero_Pedido, vc.ID_Visitante, vc.Nome_Completo, vc.Nome_Credencial, vc.CPF, vc.Email, pe.Valor_Pedido, pe.Data_Pedido, data_pagamento order by data_pagamento"
	
	
	
	
	'// Abre a conex�o com o banco e executa a string SQL
	Set objConexao = Server.CreateObject("ADODB.Connection")
	objConexao.Open Application("cnn")
	
	
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
		'// Monta as colunas de header do documento�
		Response.Write("<table border=1>")
		Response.Write("<tr>")
		Response.Write("<th>N&uacute;mero do Pedido</th>")
		Response.Write("<th>ID Visitante</th>")
		Response.Write("<th>Nome do Comprador</th>")
		Response.Write("<th>CPF</th>")
		Response.Write("<th>E-mail</th>")
		Response.Write("<th>Ingressos</th>")
		Response.Write("<th>Valor</th>")
		Response.Write("<th>Data</th>")
		Response.Write("<th>Status</th>")
		Response.Write("</tr>")
		
		'// Retorna as linhas do documento
		n_ped = ""
		tot_ped = 0
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
			Response.Write("<td nowrap>" & objRetorno("Nome_Completo") & "</td>")
			Response.Write("<td style=""mso-number-format:\@"">" & objRetorno("CPF") & "</td>")
			Response.Write("<td nowrap>" & objRetorno("Email") & "</td>")
			Response.Write("<td>" & objRetorno("Ingressos") & "</td>")
			Response.Write("<td style=""mso-number-format:'0\.00'"">" & "60"& "</td>")
			Response.Write("<td>" & objRetorno("Data_Pedido") & "</td>")
			Response.Write("<td nowrap>" & objRetorno("Status_PTB") & "</td>")
			Response.Write("</tr>")
			
			objRetorno.MoveNext()
		Wend
		
		Response.Write("</table>")
	End If
	
	'// Finaliza os objetos
	objConexao.Close
	Set objConexao = Nothing
	Set objRetorno = Nothing
	
	'// Fun��o para gerar o nome o arquivo com a hora atual
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
		
	End Function
	
	Response.Flush()
%>
