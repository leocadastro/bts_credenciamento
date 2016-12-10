<%
' Só funciona se a página tiver

' <%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%'>

'======================================================
'Converter Acento para HTML
'======================================================
Function Acentos2Htm(campo)
	limpar = campo
	limpar = Replace(limpar, "á", "&aacute;")
	limpar = Replace(limpar, "Á", "&Aacute;")
	limpar = Replace(limpar, "é", "&eacute;")
	limpar = Replace(limpar, "É", "&Eacute;")
	limpar = Replace(limpar, "í", "&iacute;")
	limpar = Replace(limpar, "Í", "&Iacute;")
	limpar = Replace(limpar, "ó", "&oacute;")
	limpar = Replace(limpar, "Ó", "&Oacute;")
	limpar = Replace(limpar, "ú", "&uacute;")
	limpar = Replace(limpar, "Ú", "&Uacute;")
	limpar = Replace(limpar, "à", "&agrave;")
	limpar = Replace(limpar, "À", "&Agrave;")
	limpar = Replace(limpar, "è", "&egrave;")
	limpar = Replace(limpar, "È", "&Egrave;")
	limpar = Replace(limpar, "ì", "&igrave;")
	limpar = Replace(limpar, "Ì", "&Igrave;")
	limpar = Replace(limpar, "ò", "&ograve;")
	limpar = Replace(limpar, "Ò", "&Ograve;")
	limpar = Replace(limpar, "ù", "&ugrave;")
	limpar = Replace(limpar, "Ù", "&Ugrave;")
	limpar = Replace(limpar, "ç", "&ccedil;")
	limpar = Replace(limpar, "Ç", "&Ccedil;")
	limpar = Replace(limpar, "â", "&acirc;")
	limpar = Replace(limpar, "Â", "&Acirc;")
	limpar = Replace(limpar, "ê", "&ecirc;")
	limpar = Replace(limpar, "Ê", "&Ecirc;")
	limpar = Replace(limpar, "î", "&icirc;")
	limpar = Replace(limpar, "Î", "&Icirc;")
	limpar = Replace(limpar, "ô", "&ocirc;")
	limpar = Replace(limpar, "Ô", "&Ocirc;")
	limpar = Replace(limpar, "û", "&ucirc;")
	limpar = Replace(limpar, "Û", "&Ucirc;")
	limpar = Replace(limpar, "ã", "&atilde;")
	limpar = Replace(limpar, "Ã", "&Atilde;")
	limpar = Replace(limpar, "õ", "&otilde;")
	limpar = Replace(limpar, "Õ", "&Otilde;")
	Acentos2Htm = limpar
End Function
'	Acentos2Htm(SQL_Video_ID)
'======================================================
'Converter Acento para HTML
'======================================================
%>