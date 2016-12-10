<!-- #include virtual="/Includes/Classes/FreeASPUpload.asp" -->
<%
    Dim DiretorioDestino
    Dim coluna, edicao

    coluna = Request.QueryString("coluna")
    edicao = Request.QueryString("edicao")

    '// Define o diretorio de destino
    Select Case coluna
        Case "faixa_fundo"
            DiretorioDestino = "/img/geral/faixa_feiras/"
        Case "logo_faixa"
            DiretorioDestino = "/img/geral/faixa_feiras/"
        Case "logo_box"
            DiretorioDestino = "/img/geral/logos/"
        Case "logo_email"
            DiretorioDestino = "/img/geral/logos/"
        Case "url_template"
            DiretorioDestino = "/template/email/"
    End Select

    If DiretorioDestino <> "" Then
        '// Faz o upload da imagem
        retornoUpload = UploadArquivo(DiretorioDestino)
    End If
    

    '// Função para gravar o arquivo enviado
    Function UploadArquivo(DiretorioDestino)
        Dim objASPUpload
        Dim retorno
        Dim Arquivos
        Dim listaArquivos

        Set objASPUpload = New FreeASPUpload
        objASPUpload.Save(Server.MapPath(DiretorioDestino))

        Arquivos = objASPUpload.UploadedFiles.keys

        If UBound(Arquivos) <> -1 Then
            For Each listaArquivos in Arquivos
                retorno = retorno & objASPUpload.UploadedFiles(listaArquivos).FileName
            Next
            UploadArquivo = DiretorioDestino & retorno
        Else
            UploadArquivo = ""
        End If

        Set objASPUpload = Nothing
    End Function
%>
<script type="text/javascript">
    if (window.opener) {
        if (window.opener.document.cad) {
            window.opener.document.cad["<%=coluna%>"].value = "<%=retornoUpload%>";
        }
    }
    window.close();
</script>
