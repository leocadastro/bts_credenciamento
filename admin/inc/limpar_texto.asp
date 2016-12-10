<%
'======================================================
' Limpando textos maliciosos
'======================================================
Function Limpar_Texto(campo)
	limpar = campo
	limpar = Replace(limpar, "'", "")
	limpar = Replace(limpar, ";", "")
	limpar = Replace(limpar, "--", "")
'	limpar = Replace(limpar, "=", "")
	limpar = Replace(limpar, """", "")
'	limpar = Replace(limpar, "and", "")
	limpar = Replace(limpar, "insert ", "")
	limpar = Replace(limpar, "update ", "")
	limpar = Replace(limpar, "select ", "")
	limpar = Replace(limpar, "delete ", "")
	limpar = Replace(limpar, "set ", "")
	limpar = Replace(limpar, "values ", "")
	Limpar_Texto = limpar
End Function
'	Limpar_Texto(SQL_Video_ID)
'======================================================
' Limpando textos maliciosos
'======================================================
%>