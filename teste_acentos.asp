
<script language=JavaScript RUNAT=SERVER>
// This function decodes the any string
// that's been encoded using URL encoding technique
function URLDecode(psEncodeString)
{
  return unescape(psEncodeString);
}
</script>
<%
response.Charset = "iso-8859-1" 

' Só funciona se a página tiver

' <%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%'>

'======================================================
'Converter Acento para HTML
'======================================================
Function Acentos(campo)
	limpar = campo
	limpar = Replace(limpar, "á", "a")
	limpar = Replace(limpar, "Á", "A")
	limpar = Replace(limpar, "é", "e")
	limpar = Replace(limpar, "É", "E")
	limpar = Replace(limpar, "í", "i")
	limpar = Replace(limpar, "Í", "I")
	limpar = Replace(limpar, "ó", "o")
	limpar = Replace(limpar, "Ó", "O")
	limpar = Replace(limpar, "ú", "u")
	limpar = Replace(limpar, "Ú", "U")
	limpar = Replace(limpar, "à", "a")
	limpar = Replace(limpar, "À", "A")
	limpar = Replace(limpar, "è", "e")
	limpar = Replace(limpar, "È", "E")
	limpar = Replace(limpar, "ì", "i")
	limpar = Replace(limpar, "Ì", "I")
	limpar = Replace(limpar, "ò", "o")
	limpar = Replace(limpar, "Ò", "O")
	limpar = Replace(limpar, "ù", "u")
	limpar = Replace(limpar, "Ù", "U")
	limpar = Replace(limpar, "ç", "c")
	limpar = Replace(limpar, "Ç", "C")
	limpar = Replace(limpar, "â", "a")
	limpar = Replace(limpar, "Â", "A")
	limpar = Replace(limpar, "ê", "e")
	limpar = Replace(limpar, "Ê", "E")
	limpar = Replace(limpar, "î", "i")
	limpar = Replace(limpar, "Î", "I")
	limpar = Replace(limpar, "ô", "o")
	limpar = Replace(limpar, "Ô", "O")
	limpar = Replace(limpar, "û", "u")
	limpar = Replace(limpar, "Û", "U")
	limpar = Replace(limpar, "ã", "a")
	limpar = Replace(limpar, "Ã", "A")
	limpar = Replace(limpar, "õ", "o")
	limpar = Replace(limpar, "Õ", "O")
	Acentos = limpar
End Function
'	Acentos2Htm(SQL_Video_ID)
'======================================================
'Converter Acento para HTML
'======================================================


response.write( Request.Form("txt") & "<br>")
response.write( URLDecode(Request.Form("txt")) & "<br>")
response.write( Acentos(URLDecode(Request.Form("txt"))))


response.write("<br>" &  Acentos("ç"))
%>
<script language="javascript" src="/js/tipos.js"></script>
<script>
	alert(isString('aaerpgo 123'));
</script>
<form method="post">
<input type="text" name="txt" />
<input type="button" value="ok" />
</form>