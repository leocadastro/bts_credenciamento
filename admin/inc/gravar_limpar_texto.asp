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
'	limpar = RemoveAcentos(limpar)
	Limpar_Texto = Ucase(Trim(limpar))
End Function
'	Limpar_Texto(SQL_Video_ID)
'======================================================
' Limpando textos maliciosos
'======================================================

'======================================================
' Remover Acentos
'======================================================
Function RemoveAcentos(ByVal Texto)
    Dim ComAcentos
    Dim SemAcentos
    Dim Resultado
	Dim Cont
    'Conjunto de Caracteres com acentos
    ComAcentos = "ΑΝΣΪΙΔΟΦάΛΐΜÒΩΘΓΥΒΞΤΫΚανσϊιδοφόλΰμςωθγυβξτϋκΗη"
    'Conjunto de Caracteres sem acentos
    SemAcentos = "AIOUEAIOUEAIOUEAOAIOUEaioueaioueaioueaoaioueCc"
    Cont = 0
    Resultado = Texto
    Do While Cont < Len(ComAcentos)
	Cont = Cont + 1
	Resultado = Replace(Resultado, Mid(ComAcentos, Cont, 1), Mid(SemAcentos, Cont, 1))
    Loop
    RemoveAcentos = Resultado
End Function
'======================================================
' Remover Acentos
'======================================================
%>