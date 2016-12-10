<%
'Instancia o objeto XMLDOM.
Set objXML = CreateObject("Microsoft.XMLDOM")
 
'Indicamos que o download em segundo plano não é permitido
objXML.async = False
 
'Carrega o domcumento XML
CaminhoArquivo = Server.MapPath("arquivos_2013\COPIA_2013-ABF.xml")

objXML.load(CaminhoArquivo)

 
'Carrega o domcumento XML
'Para quem possui serviço de REVENDA, utilize este caminho
'objXMLDoc.load("E:\vhosts\DOMINIO_COMPLETO\httpdocs\internet.xml")
 
'O método parseError contém informações sobre o último erro ocorrido
if objXML.parseError <> 0 then
 
response.write "Código do erro: " 			& objXML.parseError.errorCode & "<br>"
response.write "Posição no arquivo: " 		& objXML.parseError.filepos & "<br>"
response.write "Linha: " 					& objXML.parseError.line & "<br>"
response.write "Posição na linha: " 		& objXML.parseError.linepos & "<br>"
response.write "Descrição: " 				& objXML.parseError.reason & "<br>"
response.write "Texto que causa o erro: " 	& objXML.parseError.srcText & "<br>"
response.write "Url problemas: " 			& objXML.parseError.url
 
else
 
'A propriedade documentElement refere-se à raiz do documento
set ElemProperty 		= objXML.getElementsByTagName("credenciamento")
set ElemCadastro		= objXML.getElementsByTagName("credenciamento/cadastro")
set ElemCadastroCPF 	= objXML.getElementsByTagName("credenciamento/cadastro/cpf")
set ElemCadastroPAS		= objXML.getElementsByTagName("credenciamento/cadastro/passaporte")
 
'Looping para percorrer todos os elementos filhos
For i = 0 to (ElemProperty.length -1)
 
'A propriedade NodeName contém o nome do elemento
'A propriedade childNodes contém a lista de
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