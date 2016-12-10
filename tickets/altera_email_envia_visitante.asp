<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() -1 
Response.AddHeader "pragma", "no-store"
Response.AddHeader "cache-control","no-store, no-cache, must-revalidate"
response.Charset="ISO-8859-1"

email = request("email")
id_visitante = request("id_visitante")


Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")


sql = "update visitantes set email_envia = '" & email & "' where id_visitante=" & id_visitante


Conexao.execute(sql)


response.write "<TD colspan=3>Email alterado com sucesso</td>"


%>