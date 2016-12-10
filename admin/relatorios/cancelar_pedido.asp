<%

Response.Expires = -1
Response.ExpiresAbsolute = Now() -1 
Response.AddHeader "pragma", "no-store"
Response.AddHeader "cache-control","no-store, no-cache, must-revalidate"
response.Charset="ISO-8859-1"

id_status = request.querystring("id_status")
id_pedido = request.querystring("id_pedido")
pag = request.querystring("pag")


Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")


sql="update pedidos set status_pedido = 4 where numero_pedido = '" & id_pedido & "'"
Conexao.execute(sql)


response.redirect "relatorio_ingressos.asp?ID_Evento=5&ano_pedido=2015&id_status=" & id_status & "&pag=" & pag





%>