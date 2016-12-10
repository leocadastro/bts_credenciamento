<% @Language="VBScript" %>
<%
    Dim titulo
    Dim edicao : edicao = Request("edicao")
    Dim coluna : coluna = Request("coluna")

    '// Define o titulo da pagina
    Select Case coluna
        Case "faixa_fundo"
            titulo = "Upload - Faixa do Fundo"
        Case "logo_faixa"
            titulo = "Upload - Logotipo da Faixa"
        Case "logo_box"
            titulo = "Upload - Logotipo do Box"
        Case "logo_email"
            titulo = "Upload - Logotipo para Email de Confirm."
        Case "url_template"
            titulo = "Upload - Template do Email de Confirm."
        Case Else
            titulo = "Upload"
   End Select
    
%>
<html>  
    <head>
        <title>Admin - <%=titulo%></title>
        <style type="text/css">
            body
            {
                font-family:Arial, Helvetica, sans-serif;
                font-size: 10pt;
            }
        </style>
        <script language="javascript" type="text/javascript">
            // Valida o formulário antes de enviar as informações
            function VerificaFormulario()
            {
                var nomeArquivo;
                var extArquivo;
                var tipoPermitidos = ".gif.jpg.png";

                try {
                    nomeArquivo = document.upload.arquivo.value;
                    extArquivo = nomeArquivo.split(".");

                    if (nomeArquivo == "" || extArquivo.length <= 0) {
                        alert("Selecione o arquivo desejado!");
                        return false;
                    }
                    else if (tipoPermitidos.indexOf(extArquivo[extArquivo.length - 1].toLowerCase()) == -1) {
                        alert("Utilize apenas imagens do tipo gif, jpg ou png!");
                        return false;
                    }

                }
                catch (err) {
                    alert(err.message);
                    return false;
                }
            }
        </script>
    </head>
    <body>
        <font color="#b01d22"><h3><%=titulo%></h3></font>
        <br />
        Localize o arquivo desejado e clique em enviar.<br />
        <br />
        <br />

        <form name="upload" id="upload" action="upload2.asp?edicao=<%=edicao%>&coluna=<%=coluna%>" method="post" enctype="multipart/form-data" onsubmit="return VerificaFormulario()">
            <input type="file" name="arquivo" id="arquivo" size="50" />
            <br />
            <br />

            <input type="submit" value=" Enviar " />
        </form>

    </body>
</html>

