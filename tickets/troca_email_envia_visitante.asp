<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() -1 
Response.AddHeader "pragma", "no-store"
Response.AddHeader "cache-control","no-store, no-cache, must-revalidate"
response.Charset="ISO-8859-1"

email = request.querystring("email")
id_visitante = request.querystring("id_visitante")


Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")


Set RS_Textos = Server.CreateObject("ADODB.Recordset")

sql="select email_envia, Nome_Completo from visitantes where id_visitante = " & id_visitante
RS_Textos.Open sql, Conexao

if not RS_Textos.eof then
	
	Nome_Completo = RS_Textos(1)
	email_envia  = RS_Textos(0)

end if
RS_Textos.close


response.write "Nome:" & Nome_Completo & "<BR>" & "Email Comprovante:" & email_envia


%>