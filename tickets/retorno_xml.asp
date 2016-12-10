<!--#include virtual="/includes/limpar_texto.asp"-->
<%

URL = "https://www.aprovafacil.com/cgi-bin/STAC/informaseminarios/CAP?Transacao=" & Limpar_Texto(Request("transacao"))


' creating an object of XMLDOM
Set objXML = Server.CreateObject("Microsoft.XMLDOM")
objXML.setProperty "ServerHTTPRequest", True 
objXML.async = False

' Locating our XML database
objXML.Load(URL)
If objXML.parseError.errorCode <> 0 Then
 Response.Write "<p><font color=red>ERRO</font></p>"
 Response.End
End If


Set objLst = objXML.getElementsByTagName("ResultadoCAP")

For i = 0 To objLst.Length - 1 


 Set subLst = objLst.item(i)
 Response.Write subLst.childNodes(0).childNodes(0).Text
Next 


%>