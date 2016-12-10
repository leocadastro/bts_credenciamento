<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Response.Charset="ISO-8859-1"%>
<!--#include virtual="/includes/limpar_texto.asp"-->
<!--#include virtual="/scripts/ConsultarWebService.asp"-->
<!-- #include file ="paypalfunctions.asp" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content-type:"text/html; charset=ISO-8859-1" />
<title>Credenciamento <%=Year(Now())%> - BTS Informa</title>
<link href="http://credenciamento.btsinforma.com.br/css/base_forms.css" rel="stylesheet" type="text/css" />
<link href="http://credenciamento.btsinforma.com.br/css/estilos.css" rel="stylesheet" type="text/css">

<%

dim b

Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
Set RS_Consulta_Pedido = Server.CreateObject("ADODB.Recordset")

sql="select visitantes.cpf, visitantes.email  from  Pedidos_Historico inner join pedidos on pedidos.Numero_Pedido = Pedidos_Historico.Numero_Pedido inner join  Pedidos_Carrinho on Pedidos_Carrinho.ID_Pedido = pedidos.ID_Pedido inner join Visitantes on visitantes.ID_Visitante = Pedidos_Carrinho.id_visitante where  Pedidos_Historico.Data_Pagamento like '%2015%' and (Pedidos_Historico.codigo_autorizacao like 'SUCCESS' or Pedidos_Historico.codigo_autorizacao like 'SUCCESSWITHWARNING') and Pedidos_Historico.data_pagamento > '2015-04-29 16:00:00.000' and Pedidos_Historico.Numero_Pedido <> '' and Pedidos_Historico.Status_Pagamento = 1 and visitantes.email not like '%paypal%'"


RS_Consulta_Pedido.Open sql, Conexao, 3, 3

base=""
i=0
do while not RS_Consulta_Pedido.eof
i = i +1
base = RS_Consulta_Pedido(0)
'response.write trim(base)&"<BR>"
'response.write len(trim(base))&"<BR>"
'response.end

if len(trim(base)) < 11 then base = RS_Consulta_Pedido(1)


if base <> "04636744993" and base <> "04318526933" and base <> "41210155842" then 
 b = SetComprador(base)

end if

'response.write i & ": " & base & "<BR>"
		
base=""


RS_Consulta_Pedido.movenext
loop

RS_Consulta_Pedido.close
response.write "foi " & i

%>