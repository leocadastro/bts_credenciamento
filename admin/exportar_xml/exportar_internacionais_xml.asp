<%
'Instancia o objeto XMLDOM.
Set objXML = CreateObject("Microsoft.XMLDOM")
 
'Indicamos que o download em segundo plano n�o � permitido
objXML.async = False
 
'Carrega o domcumento XML
CaminhoArquivo = Server.MapPath("arquivos_2013\COPIA_2013-ABF.xml")

objXML.load(CaminhoArquivo)

 
'Carrega o domcumento XML
'Para quem possui servi�o de REVENDA, utilize este caminho
'objXMLDoc.load("E:\vhosts\DOMINIO_COMPLETO\httpdocs\internet.xml")
 
'O m�todo parseError cont�m informa��es sobre o �ltimo erro ocorrido
if objXML.parseError <> 0 then
 
response.write "C�digo do erro: " 			& objXML.parseError.errorCode & "<br>"
response.write "Posi��o no arquivo: " 		& objXML.parseError.filepos & "<br>"
response.write "Linha: " 					& objXML.parseError.line & "<br>"
response.write "Posi��o na linha: " 		& objXML.parseError.linepos & "<br>"
response.write "Descri��o: " 				& objXML.parseError.reason & "<br>"
response.write "Texto que causa o erro: " 	& objXML.parseError.srcText & "<br>"
response.write "Url problemas: " 			& objXML.parseError.url
 
else
 
'A propriedade documentElement refere-se � raiz do documento
set ElemProperty 		= objXML.getElementsByTagName("credenciamento")
set ElemCadastro		= objXML.getElementsByTagName("credenciamento/cadastro")
set ElemCadastroCPF 	= objXML.getElementsByTagName("credenciamento/cadastro/cpf")
set ElemCadastroPAS		= objXML.getElementsByTagName("credenciamento/cadastro/passaporte")
 
'Looping para percorrer todos os elementos filhos
For i = 0 to (ElemProperty.length -1)
 
'A propriedade NodeName cont�m o nome do elemento
'A propriedade childNodes cont�m a lista de
'elementos filhos

	'response.Write(raiz.ChildNodeName)
	
	response.Write("CPF: " & ElemCadastroCPF & "<br>")
	
	'If raiz.NodeName.item(20) = "Passaporte" Then
	'	response.write()
	'End If 

	'Response.write raiz.NodeName & "<br>" & raiz.childNodes.item(i).childNodes.item(20).text & "<br>" &  raiz.childNodes.item(i).childNodes.item(21).text
 
Next
 
end if
 
'Destruindo os objetos
Set objXMLDoc = Nothing
Set raiz = Nothing
%>