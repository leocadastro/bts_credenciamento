<%
'======================================================
' Limpando caracteres especiais
'======================================================
Function Limpar_Texto(campo)
	limpar = campo
	limpar = Replace(limpar, "'", "''")
	limpar = Replace(limpar, "--", "")
	limpar = Replace(limpar, "&lt;", "<")
	limpar = Replace(limpar, "&gt;", ">")	
	limpar = Replace(limpar, "&amp;", "&")
	limpar = Replace(limpar, "&nbsp;", " ")
	limpar = Replace(limpar, "&ndash;", "�")
	limpar = Replace(limpar, "&mdash;", "�")
	limpar = Replace(limpar, "&hellip;", "�")
	limpar = Replace(limpar, "&bull;", "&bull;")
	limpar = Replace(limpar, "&sect;", "&sect;")
	limpar = Replace(limpar, "&copy;", "�")
	limpar = Replace(limpar, "&reg", "�")
	limpar = Replace(limpar, "&ordf;", "�")
	
	
	limpar = Replace(limpar, "&aacute;", "�") 
	limpar = Replace(limpar, "&Aacute;", "�") 
	limpar = Replace(limpar, "&atilde;", "�") 
	limpar = Replace(limpar, "&Atilde;", "�") 
	limpar = Replace(limpar, "&acirc;", "�")
	limpar = Replace(limpar, "&Acirc;", "�")
	limpar = Replace(limpar, "&agrave;", "�")
	limpar = Replace(limpar, "&Agrave;", "�") 
	limpar = Replace(limpar, "&eacute;", "�")
	limpar = Replace(limpar, "&Eacute;", "�")
	limpar = Replace(limpar, "&ecirc;", "�")
	limpar = Replace(limpar, "&Ecirc;", "�") 
	limpar = Replace(limpar, "&iacute;", "�")
	limpar = Replace(limpar, "&Iacute;", "�")
	limpar = Replace(limpar, "&oacute;", "�")
	limpar = Replace(limpar, "&Oacute;", "�") 
	limpar = Replace(limpar, "&otilde;", "�") 
	limpar = Replace(limpar, "&Otilde;", "�") 
	limpar = Replace(limpar, "&ocirc;", "�") 
	limpar = Replace(limpar, "&Ocirc;", "�") 
	limpar = Replace(limpar, "&uacute;", "�")
	limpar = Replace(limpar, "&Uacute;", "�")
	limpar = Replace(limpar, "&uuml;", "�")
	limpar = Replace(limpar, "&Uuml;", "�") 
	limpar = Replace(limpar, "&ccedil;", "�")
	limpar = Replace(limpar, "&Ccedil;", "�") 
	limpar = Replace(limpar, "&ntilde;", "�") 	
	limpar = Replace(limpar, "&Ntilde;", "�") 	
	limpar = Replace(limpar, "&quot;", """")
	limpar = Replace(limpar, "insert ", "")
	limpar = Replace(limpar, "update ", "")
	limpar = Replace(limpar, "select ", "")
	limpar = Replace(limpar, "delete ", "")
	limpar = Replace(limpar, "set ", "")
	limpar = Replace(limpar, "values ", "")
	Limpar_Texto = limpar
End Function
%>