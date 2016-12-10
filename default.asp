<!--#include virtual="/includes/limpar_texto.asp"-->
<%
' Verificando Acesso Externo
Out = limpar_texto(Request("Out"))
If Out <> "" Then
	Acesso = AcessoExterno(Out)
End If

Session.Timeout = 240

'===========================================================
Qs = Request.ServerVariables("QUERY_STRING")
'===========================================================

possibilitar_troca_idioma = true

If Session("cliente_idioma") = "" AND Request("i") = "" Then 
	Session("cliente_idioma") = 1 ' Portugues
Else 
	i = Limpar_Texto(Request("i"))
	If  Len(i) = 1 AND isNumeric(i) = True Then	Session("cliente_idioma") = i
End If
idioma = Session("cliente_idioma")
'===========================================================
' Limpar Sessões do CLIENTE
Dim item
For Each item in Session.Contents
	If (Left(item,7) = "cliente") and item <> "cliente_msg" Then
		Session(item) = ""
	End If
Next
Session("cliente_logado") = false
'===========================================================	
	Set Conexao = Server.CreateObject("ADODB.Connection")
	Conexao.Open Application("cnn")

	Pagina_ID = 1
	
	SQL_Textos	=	" Select " &_
					"	ID_Texto, " &_
					"	ID_Tipo, " &_
					"	Identificacao, " &_
					"	Texto, " &_
					"	URL_Imagem " &_
					" From Paginas_Textos " &_
					" Where  " &_
					"	ID_Idioma = " & idioma &_
					"	AND ID_Pagina = " & Pagina_ID &_
					" Order By ID_Texto "

	Set RS_Textos = Server.CreateObject("ADODB.Recordset")
	RS_Textos.Open SQL_Textos, Conexao
	
	If not RS_Textos.BOF or not RS_Textos.EOF Then
		total_registros = 0
		While not RS_Textos.EOF
			total_registros = total_registros + 1
			RS_Textos.MoveNext
		Wend
		RS_Textos.MoveFirst
		ReDim textos_array(total_registros-1)
		n = 0
		While not RS_Textos.EOF
			id = RS_Textos("id_texto")
			ident = RS_Textos("identificacao")
			texto = RS_Textos("texto")
			url_img = RS_Textos("url_imagem")
			textos_array(n) = Array(id, ident, texto, url_img)
			n = n + 1
			RS_Textos.MoveNext
		Wend
		RS_Textos.Close
	End If

'If Request("teste") = "sim" Then
'	For i = Lbound(textos_array) to Ubound(textos_array)
'		response.write("[ i: " & i & " ] [ ident: " & textos_array(i)(1) & " ]  [ txt: " & textos_array(i)(2) & " ]  [ img: " & textos_array(i)(3) & " ]<br>")
'	Next
'End If
'===========================================================
	' Listagem de Feiras por DATA
	SQL_Feiras	= 	"Select " &_
					"	Distinct " &_
					"	Ee.ID_Edicao, " &_
					"	Ecv.Cor, " &_
					"	Ecv.Logo_Box, " &_
					"	Ecv.Logo_Negativo, " &_
					"	Ecv.Faixa_Fundo, " &_
					"	E.Nome_PTB as Evento, " &_
					"	Ee.Ano, " &_
					"	Ee.Data_Inicio_Feira " &_
					"From Edicoes_Configuracao as Ecv  " &_
					"Inner Join Eventos_Edicoes as Ee ON Ee.ID_Edicao = Ecv.ID_Edicao  " &_
					"Inner Join Eventos as E ON E.ID_Evento = Ee.ID_Evento  " &_
					"Inner Join Edicoes_Tipo as Et ON Et.ID_Edicao = Ecv.ID_Edicao " &_
					"Where " &_
					"	Ecv.Ativo = 1 " &_
					"	AND Ee.Ativo = 1 " &_
					"	AND E.Ativo = 1 " &_
					"	AND Et.Ativo = 1 " &_
					"	AND getDate() >= Et.Inicio " &_
					"	AND getDate() <= Et.Fim " &_
					"Order by Ee.Data_Inicio_Feira, Evento "
' ! ATENÇÃO ==============================================================
					' Alteração pro ambiente de TESTE
'					"	AND getDate() >= Et.Inicio " &_
'					"	AND getDate() <= Et.Fim " &_
' ! ATENÇÃO ==============================================================

	
'response.write("<b>SQL_Feiras</b><br>" & SQL_Feiras & "<hr>")
	
	Set RS_Feiras = Server.CreateObject("ADODB.Recordset")
	RS_Feiras.CursorType = 0
	RS_Feiras.LockType = 1
	RS_Feiras.Open SQL_Feiras, Conexao, 1
'===========================================================
%>
<% If Request("teste") = "s" Then %>
	<!--#include virtual="/includes/exibir_array.asp"-->
<% End IF

If Request("teste") = "sim" Then
	teste = "sim"
Else
	teste = "não"
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Credenciamento - BTS Informa</title>
<link href="/css/estilos.css" rel="stylesheet" type="text/css">
<link href="/css/jquery.alerts.css" rel="stylesheet" type="text/css">
<script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
<script language="javascript" src="/js/jquery.alerts.js"></script>
<script language="javascript" src="/js/jquery-ui-1.8.7.core_eff-slide.js"></script>
<script language="javascript" src="/js/funcoes_gerais.js"></script>
<!-- "/includes/google_analytics.asp" -->
<script language="javascript" src="/scripts/default.js?v=2"></script>
<script language="javascript">
	// Definição das Feiras
	<%
	If not RS_Feiras.BOF or not RS_Feiras.EOF Then
		i = 0
		While not RS_Feiras.EOF
			cor 		= RS_Feiras("cor")
			logo		= RS_Feiras("logo_box")
			faixa_fundo	= RS_Feiras("faixa_fundo")
			faixa_logo	= RS_Feiras("logo_negativo")
			id_edicao	= RS_Feiras("ID_Edicao")
			i = i + 1
			
			' Projeto 0014 - Item 3
			' Pegar o ID da Configuracao da feira escolhida
			If Len(Trim(Request.QueryString("f"))) > 0 Then
				If Cstr(Limpar_Texto(Request.QueryString("f"))) = Cstr(id_edicao) Then 
					link_configuracao = i
				End If
			End If
			
			%><%=VBLF%>configuracao_feira[<%=i%>] = Array('<%=cor%>','<%=logo%>','<%=faixa_fundo%>','<%=faixa_logo%>','<%=id_edicao%>');<%
			RS_Feiras.MoveNext
		Wend
		RS_Feiras.MoveFirst()
	End If
	%>
	$(document).ready(function(){
		// Aviso
		var erro = '<%=Request("erro")%>';
		switch (erro) {
			case '1':
				jAlert('Sua sess&atilde;o expirou, inicie novamente.','Aten&ccedil;&atilde;o');
				break;	
			default:
				break;
		}
<%
'===========================================================
	'*********************************************
	' Implementação Projeto 0014 - Item 3
	' Data: 29 / Nov / 2012
	' Por: Homero
	' Descrição: Link Externo para Feira e Idioma
	'*********************************************

	If Len(Trim(Request.QueryString("f"))) > 0 Then
		' Idioma
		link_idioma = Cint(Limpar_Texto(Request.QueryString("i")))
		' Código Feira
		link_feira = Cint(Limpar_Texto(Request.QueryString("f")))
		
		' Verificacao de Segurança
		If IsNumeric(link_idioma) AND IsNumeric(link_feira) Then
			response.write("idioma(" & link_idioma & ");")
			response.write("proxima_tela(" & link_configuracao & "," & link_feira & ");")
		End If
	End If
'===========================================================
%>		
	});
	var tp_formulario = '';
	
	
  (function(i,s,o,g,r,a,m){i['GoogleAnalyticsObject']=r;i[r]=i[r]||function(){
  (i[r].q=i[r].q||[]).push(arguments)},i[r].l=1*new Date();a=s.createElement(o),
  m=s.getElementsByTagName(o)[0];a.async=1;a.src=g;m.parentNode.insertBefore(a,m)
  })(window,document,'script','//www.google-analytics.com/analytics.js','ga');

  ga('create', 'UA-21195903-15', 'btsinforma.com.br');
  ga('send', 'pageview');
</script>
</head>

<body id="conteudo">
<div id='pre-load' style="display:none; visibility:hidden;">
<%
	If not RS_Feiras.BOF or not RS_Feiras.EOF Then
		i = 0
		While not RS_Feiras.EOF
			logo		= RS_Feiras("logo_box")
			faixa_fundo	= RS_Feiras("faixa_fundo")
			faixa_logo	= RS_Feiras("logo_negativo")
			%><img src="<%=logo%>"><%=VBLF%><%
			%><img src="<%=faixa_fundo%>"><%=VBLF%><%
			%><img src="<%=faixa_logo%>"><%=VBLF%><%
			RS_Feiras.MoveNext
		Wend
		RS_Feiras.MoveFirst()
	End If
%>
</div>
<!--#include virtual="/includes/cabecalho.asp"-->
<div style="width: 100%; position: absolute; left:0px; float:left; z-index:10;" id="faixa">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="33%" align="center" height="45">
    <!-- Faixa Lateral -->
    	<div style="background:url(/img/geral/faixa_fundo_esq.gif); height:45px; width:100%; margin-top:50px;"></div>
    <!-- Faixa Lateral -->
    </td>
    <td width="870" align="center">
    <!-- Faixa -->
        <table width="870" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr>
            <td height="50">&nbsp;</td>
          </tr>
          <tr>
            <td>
                <table width="870" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="189" height="45" background="/img/geral/faixa_fundo_esq.gif"><img id="img_faixa_esq" src="/img/geral/tipos/faixa_visitantes.gif" width="189" height="45" title="" alt=""></td>
                    <td height="45" background="/img/geral/faixa_fundo_dir.gif" class="atencao_13px cor_branco">
                   	  <div id="txt_1" style="padding-left:20px; float:left; height:45px; line-height:40px;" align="left">Escolha a feira em que você deseja se Credenciar</div>
                    </td>
                  </tr>
                </table>
            </td>
          </tr>
      </table>
    <!-- Faixa -->
    </td>
    <td width="33%" align="center" valign="top">
    <!-- Faixa Lateral -->
    	<div style="background:url(/img/geral/faixa_fundo_dir.gif); height:45px; width:100%; margin-top:50px;"></div>
    <!-- Faixa Lateral	 -->
    </td>
  </tr>
</table>
</div>
<div style="width: 100%; position: absolute; left:0px; float:left; z-index:10; " id="faixa_selecionada">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="33%" align="center">
    <!-- Faixa Lateral -->
    	<div style="background:url(/img/geral/spacer.gif); height:45px; width:100%; margin-top:50px;"></div>
    <!-- Faixa Lateral -->
    </td>
    <td width="870" align="center">
    <!-- Faixa -->
        <table width="870" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr>
            <td height="50">&nbsp;</td>
          </tr>
          <tr>
            <td>
                <table width="870" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="189" height="45">&nbsp;</td>
                    <td id="img_fundo_selecionado" height="45" background="/img/geral/faixa_fundo_dir.gif" class="atencao_13px cor_branco">
                    	<div id="txt_2" style="padding-left:20px; float:left; line-height:40px;" align="left">Qual o tipo de Credencial desejada ?</div>
                        <div style="float:right;" align="right"><img id="img_logo_selecionado" src="/img/geral/faixa_fundo_dir.gif" hspace="10"></div>
                    </td>
                  </tr>
                </table>
            </td>
          </tr>
      </table>
    <!-- Faixa -->
    </td>
    <td width="33%" align="center" valign="top">
    <!-- Faixa Lateral -->
    	<div style="background:url(/img/geral/spacer.gif); height:45px; width:100%; margin-top:50px;" id="faixa_dir"></div>
    <!-- Faixa Lateral	 -->
    </td>
  </tr>
  <tr>
    <td align="center">&nbsp;</td>
    <td align="right"><a href="javascript:voltar();"><img src="/img/geral/icones/voltar.gif" width="52" height="15" border="0"></a></td>
    <td align="center" valign="top">&nbsp;</td>
  </tr>
</table>
</div>
<div style="width: 100%; position: absolute; left:0px; float:left; " id="feiras">
<table width="870" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="150" colspan="4">&nbsp;</td>
  </tr>
	<%
	If not RS_Feiras.BOF or not RS_Feiras.EOF Then
		colunas = 1
		i = 0
		While not RS_Feiras.EOF
			cor 		= RS_Feiras("cor")
			logo		= RS_Feiras("logo_box")
			evento		= RS_Feiras("evento") & " - " & RS_Feiras("ano")
			ID_Edicao	= RS_Feiras("ID_Edicao")
			data		= FormatDateTime(RS_Feiras("Data_Inicio_Feira"),2)
			i = i + 1
			If colunas = 1 Then
				%><tr><%=vbCr%><%
			End If

			If Cstr(ID_Edicao) = "21" Then
			%>
			  	<td align="center">
				<table width="260" border="0" cellspacing="0" cellpadding="0" class="bg_feira cursor" onClick="window.open('http://www.vitafoodssouthamerica.com.br/btscre/');" title="<%=evento%>" align="center">
				  <tr>
					<td height="4" bgcolor="<%=cor%>"><img src="/img/geral/spacer.gif" width="110" height="4"></td>
				  </tr>
				  <tr>
					<td height="95" align="center" class="verdana fs_10px cor_cinza2"><img src="<%=logo%>" border="0" title="<%=evento%>" alt="<%=evento%>"></td>
                    <!-- <br><br><span id="tit_inicio<%=i%>">In&iacute;cio da Feira:</span>&nbsp;<%=data%> -->
				  </tr>
				  <tr>
					<td height="4" bgcolor="<%=cor%>"><img src="/img/geral/spacer.gif" width="110" height="4"></td>
				  </tr>
				</table>
                <span class="verdana fs_10px cor_cinza2"></span>
			  </td>
			<%
			Else
			%>
			<td align="center">
				<table width="260" border="0" cellspacing="0" cellpadding="0" class="bg_feira cursor" onClick="proxima_tela('<%=i%>','<%=ID_Edicao%>','<%=teste%>');" title="<%=evento%>" align="center">
				  <tr>
					<td height="4" bgcolor="<%=cor%>"><img src="/img/geral/spacer.gif" width="110" height="4"></td>
				  </tr>
				  <tr>
					<td height="95" align="center" class="verdana fs_10px cor_cinza2"><img src="<%=logo%>" border="0" title="<%=evento%>" alt="<%=evento%>"></td>
                    <!-- <br><br><span id="tit_inicio<%=i%>">In&iacute;cio da Feira:</span>&nbsp;<%=data%> -->
				  </tr>
				  <tr>
					<td height="4" bgcolor="<%=cor%>"><img src="/img/geral/spacer.gif" width="110" height="4"></td>
				  </tr>
				</table>
                <span class="verdana fs_10px cor_cinza2"></span>
			  </td>
			<%
			End If

			RS_Feiras.MoveNext
			colunas = colunas + 1
			If RS_Feiras.EOF and colunas < 4 Then
				%>
                </tr><%=vbCr%>
				<%
			ElseIf colunas = 4 Then
				%>
                </tr><%=vbCr%>
                <tr>
                  <td width="547" height="50" colspan="4">&nbsp;</td>
                </tr>
				<%
				colunas = 1
			End If
		Wend
		RS_Feiras.Close
	End If
	%>
    <script language="javascript">
	var total_feiras = <%=i%>;
	</script>
    <tr>
      <td width="547" height="50" colspan="3">&nbsp;</td>
    </tr>
</table>
</div>
<div style="width: 100%; position: absolute; left:0px; float:left; display:none;" id="tipos">

</div>
<div style="width: 100%; position: absolute;float:left; display:none; z-index:100" id="loading">
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center"><img src="/img/geral/ico_ajax-loader.gif" style="opacity:100"></td>
  </tr>
</table>
</div>
<form action="definir.asp" method="post" id="continuar" name="continuar">
	<input type="hidden" id="idioma" name="idioma" value="">
	<input type="hidden" id="edicao" name="edicao" value="">
	<input type="hidden" id="tipo" name="tipo" value="">
	<input type="hidden" id="formulario" name="formulario" value="">
	<input type="hidden" id="url" name="url" value="">
</form>
</body>
</html>