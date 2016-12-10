<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%

'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================


paises = "AC|Acre/AL|Alagoas/AM|Amazonas/AP|Amapá/BA|Bahia/CE|Ceará/DF|Distrito Federal/ES|Espírito Santo/GO|Goiás/MA|Maranhão/MT|Mato Grosso/MS|Mato Grosso do Sul/MG|Minas Gerais/PA|Pará/PB|Paraíba/PR|Paraná/PE|Pernambuco/PI|Piauí/RJ|Rio de Janeiro/RN|Rio Grande do Norte/RO|Rondônia/RS|Rio Grande do Sul/RR|Roraima/SC|Santa Catarina/SE|Sergipe/SP|São Paulo/TO|Tocantins"

array_paises = split(paises,"/")

For i = Lbound(array_paises) To Ubound(array_paises)

	'response.write(array_paises(i) & "<br>")

	array_siglas = split(array_paises(i),"|")

	For x = Lbound(array_siglas) To Ubound(array_siglas)-1
		sigla = array_siglas(x)
		estado = array_siglas(x+1)


	response.write("- sigla: " & sigla & "<br>")
	response.write("- estado: " & estado & "<hr>")

		
		'Insert into SQL 
		'array_paises(i)
		SQL_Estado = 	"INSERT INTO UF " &_
						"	(Estado, Sigla) " &_
						"VALUES  " &_
						"	('" & estado & "', '" & sigla & "')"
				
		Set RS_Estado = Server.CreateObject("ADODB.Recordset")
		RS_Estado.Open SQL_Estado, Conexao
	
	
	Next
	
Next

Conexao.Close
%>