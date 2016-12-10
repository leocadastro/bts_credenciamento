<%
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")

Conexao.Execute ("COMMIT TRANSACTION; ")
Conexao.Close
%>