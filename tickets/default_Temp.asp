<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/admin/inc/limpar_texto.asp"-->
<!--#include virtual="/includes/texto_caixaAltaBaixa.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Credenciamento <%=Year(Now())%> - BTS Informa</title>
<link href="/css/base_forms.css" rel="stylesheet" type="text/css" />
<link href="/css/estilos.css" rel="stylesheet" type="text/css">
<link href="/css/jquery.alerts.css" rel="stylesheet" type="text/css">
<link href="/css/checkbox.css" rel="stylesheet" type="text/css">
<script language="javascript" src="/js/jquery-1.3.2.min.js"></script>
<script language="javascript" src="/js/jquery-ui-1.8.7.core_eff-slide.js"></script>
<script language="javascript" src="/js/jquery.alerts.js"></script>
<script language="javascript" src="/js/jquery.screwdefaultbuttons.js"></script>
<script language="javascript" src="/js/validar_forms.js"></script>	
<script language="javascript" src="/js/funcoes_gerais.js"></script>
<script language="javascript" src="/js/tipos.js"></script>
<!-- Script desta página -->
<script language="javascript" src="default.js" charset="utf-8"></script>
<!-- Script desta página FIM -->
<%
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================

Session("cliente_edicao") 		= 22
Session("cliente_idioma") 		= 1
Session("cliente_tipo") 		= 1
Session("cliente_formulario")	= 1

If Session("cliente_edicao") = "" OR Session("cliente_idioma") = "" OR Session("cliente_tipo") = "" Then
	response.Redirect("/?erro=1")
End If

ID_Edicao 				= Session("cliente_edicao")
Idioma 					= Session("cliente_idioma")
ID_TP_Credenciamento 	= Session("cliente_tipo")
TP_Formulario 			= Session("cliente_formulario")

	' Verifica Idioma a ser apresentado
	Select Case (Idioma)
		Case "1"
			SgIdioma = "PTB"
		Case "2"
			SgIdioma = "ESP"
		Case "3"
			SgIdioma = "ENG"
		Case Else
			SgIdioma = "PTB"
	End Select

	Pagina_ID = 2

	SQL_Textos	=	" Select " &_
					"	ID_Texto, " &_
					"	ID_Tipo, " &_
					"	Identificacao, " &_
					"	Texto, " &_
					"	URL_Imagem " &_
					" From Paginas_Textos " &_
					" Where  " &_
					"	ID_Idioma = " & idioma & " " &_
					"	AND ID_Pagina = " & Pagina_ID & " " &_
					" Order By Ordem "
	'response.write(SQL_Textos)
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

	
'	For i = Lbound(textos_array) to Ubound(textos_array)
'		response.write("[ i: " & i & " ] [ ident: " & textos_array(i)(1) & " ]  [ txt: " & textos_array(i)(2) & " ]  [ img: " & textos_array(i)(3) & " ]<br>")
'	Next
'===========================================================
%>

<script language="javascript">
	var idioma_atual = '<%=Session("cliente_idioma")%>';
	var select       = '<%=textos_array(36)(2)%>';
	var cor_fundo 	 = '<%=faixa_cor%>';
	var tp_formulario = '';
</script>


<% If Request("teste") = "s" Then %>
	<!--#include virtual="/includes/exibir_array.asp"-->
<% End IF
	
	' Select IMG Faixa
	SQL_Img_Faixa 	=	"Select " &_
						"	Img_Faixa " &_
						"From Tipo_Credenciamento " &_
						"Where ID_Tipo_Credenciamento = " & ID_TP_Credenciamento
	'response.write(SQL_Img_Faixa)
	Set RS_Img_Faixa = Server.CreateObject("ADODB.Recordset")
	RS_Img_Faixa.CursorType = 0
	RS_Img_Faixa.LockType = 1
	RS_Img_Faixa.Open SQL_Img_Faixa, Conexao
		img_faixa = RS_Img_Faixa("img_faixa")
	RS_Img_Faixa.Close
	
	' Faixa TOPO
	SQL_Faixa	= 	"Select " &_
					"	Cor, " &_
					"	Logo_Negativo, " &_
					"	Faixa_Fundo " &_
					"From Edicoes_configuracao " &_
					"Where  " &_
					"	ID_Edicao = " & ID_Edicao
	Set RS_Faixa = Server.CreateObject("ADODB.Recordset")
	RS_Faixa.CursorType = 0
	RS_Faixa.LockType = 1
	RS_Faixa.Open SQL_Faixa, Conexao
		
		faixa_cor	= RS_Faixa("cor")
		faixa_logo	= RS_Faixa("logo_negativo")
		faixa_fundo	= RS_Faixa("Faixa_Fundo")
	RS_Faixa.Close
	
	' Select de Eventos
	SQL_Evento	=	"SELECT " &_
					"	Nome_" & SgIdioma & " AS Evento, " &_
					"	Ano " &_
					"FROM Eventos as E " &_
					"INNER JOIN" &_
					"	Eventos_Edicoes as EE " &_
					"ON EE.ID_Evento = E.ID_Evento " &_
					"WHERE " &_
					"	E.Ativo = 1 " &_ 
					"	AND EE.ID_Edicao = " & ID_Edicao 

	Set RS_Evento = Server.CreateObject("ADODB.Recordset")
	RS_Evento.CursorType = 0
	RS_Evento.LockType = 1
	RS_Evento.Open SQL_Evento, Conexao
	
	Evento = RS_Evento("Evento") & " " & RS_Evento("Ano")
	Rs_Evento.Close
	
' Validar CNPJ E SENHA
'==================================================
erro = ""
If Limpar_Texto(Request.Form("frmID_Visitante")) <> "" AND Limpar_Texto(Request.Form("frmSenha")) <> "" Then

	ID_Visitante	= Limpar_Texto(Request.Form("frmID_Visitante"))
	Senha			= Limpar_Texto(Request.Form("frmSenha"))
	
	ID_Visitante	= Replace(ID_Visitante,".","")
	ID_Visitante	= Replace(ID_Visitante,"-","")

	'response.write(ID_Visitante)
	'response.write("<br>" & Senha)

	' Se vier vazio	
	If Len(Trim(ID_Visitante)) = 0 OR Len(Trim(Senha)) = 0 Then
		erro = "1"
	Else
		' Verificar se existe Cadastro
		SQL_Cadastro_Visitantes =	"Select " &_
									"	RC.ID_Relacionamento_Cadastro as IRC " &_
									"	,RC.ID_Edicao " &_
									"	,RC.ID_Empresa " &_
									"	,RC.ID_Visitante " &_
									"	,V.Nome_Completo " &_
									"	,V.CPF " &_
									"	,V.Senha " &_
									"From Relacionamento_Cadastro as RC " &_
									"Inner Join Tipo_Credenciamento as TC ON TC.ID_Tipo_Credenciamento = RC.ID_Tipo_Credenciamento " &_
									"Inner Join Visitantes as V ON V.ID_Visitante = RC.ID_Visitante " &_
									"Where  " &_
									"	RC.ID_Tipo_Credenciamento in (1,10,11,12)	/* Pessoa Fisica PTB, ESP e ENG	*/ " &_
									"	AND TC.ID_Idioma = " & Idioma & "		/* Idioma	*/ " &_
									"	AND " &_
									"	( " &_
									"	V.ID_Visitante = '" & ID_Visitante & "'  " &_
									"	or V.CPF = '" & ID_Visitante & "' " &_
									"	or V.Passaporte = '" & ID_Visitante & "' " &_
									"	)  " &_
									"	/* ID_Relacioanmento_Cadastro	*/ " &_
									"	AND V.Senha = '" & Senha & "' /* Senha	*/ " &_
									"Order by V.Data_Atualizacao DESC, V.Data_Cadastro DESC, RC.ID_Edicao DESC "
	
		'response.write(SQL_Cadastro_Visitantes)
		'response.write("<br><br>" & Session("cliente_edicao"))
		'response.End()
	
		Set RS_Cadastro_Visitantes = Server.CreateObject("ADODB.Recordset")
		RS_Cadastro_Visitantes.Open SQL_Cadastro_Visitantes, Conexao
		
		If RS_Cadastro_Visitantes.BOF or RS_Cadastro_Visitantes.EOF Then
			erro = "2"
		Else
		
			Do While Not RS_Cadastro_Visitantes.Eof
			
				If Cint(RS_Cadastro_Visitantes("ID_Edicao")) = Cint(Session("cliente_edicao")) Then
					'Session("IRC") 			= RS_Cadastro_Visitantes("IRC")
					'Session("ID_Empresa") 		= RS_Cadastro_Visitantes("ID_Empresa")
					'Session("ID_Visitante") 	= RS_Cadastro_Visitantes("ID_Visitante")
					'Session("Nome_Completo") 	= RS_Cadastro_Visitantes("Nome_Completo")
					'Session("CPF") 			= RS_Cadastro_Visitantes("CPF")

					%>
					<form id="visitantes" name="visitantes" method="POST" action="status.asp">
						<input type="hidden" name="id_tipo" 			value="13">
						<input type="hidden" name="frmID_Cadastro" 		value="<%=RS_Cadastro_Visitantes("IRC")%>">
						<input type="hidden" name="frmID_Empresa" 		value="<%=RS_Cadastro_Visitantes("ID_Empresa")%>">
						<input type="hidden" name="frmID_Visitante" 	value="<%=RS_Cadastro_Visitantes("ID_Visitante")%>">
						<input type="hidden" name="frmNome" 			value="<%=RS_Cadastro_Visitantes("Nome_Completo")%>">
						<input type="hidden" name="frmCPF" 				value="<%=RS_Cadastro_Visitantes("CPF")%>">
					</form>
					
					
					<script language="Javascript">
						/*function Enviar(){
							document.forms['visitantes'].submit();
							//window.location.href("status.asp");
						}*/
						/*
						$(document).ready(function(){
							$("#visitantes").submit();
		
						});
						*/
							document.forms['visitantes'].submit();
					</script>
					<%
					Exit Do
				Else
					erro = "3"
				End If
			
			RS_Cadastro_Visitantes.MoveNext
			Loop	
		
		RS_Cadastro_Visitantes.Close
		End If
	End If
End If

If erro <> "" Then
'response.writE("<br>erro: " & erro)
%>
	<script language="javascript">
	$(document).ready(function(){
		var erro = '<%=erro%>';
		switch (erro) {
			case '1':
				$('#txt_topo').html('Campos Vazios');
				$('#aviso_topo').show();
				break;
			case '2':
				$('#txt_topo').html('Dados incorretos');
				$('#aviso_topo').show();
				break;
			case '3':
				$('#txt_topo').html('Ainda não possui cadastro para esta edição!');
				$('#aviso_topo').show();
				break;
		}
	});
	</script>
<% 
End If 
%>
</head>

<body>
<div style="width: 100%; position: absolute; height:750px; float:left; z-index:100; background:#CCC; display:none" id="loading" class="transparent">
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center"><img src="/img/geral/ico_ajax-loader.gif" style="opacity:100"></td>
  </tr>
</table>
</div>
<!--#include virtual="/includes/cabecalho.asp"-->
<div style="width: 100%; position: absolute; left:0px; float:left; z-index:10; height: 115px;" id="faixa_selecionada">
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
                    <td width="189" height="45" background="/img/geral/faixa_fundo_esq.gif"><img id="img_faixa_esq" src="/img/geral/tipos/Faixa_Tickets.gif" width="189" height="45"></td>
                    <td id="img_fundo_selecionado" height="45" background="<%=faixa_fundo%>" class="atencao_13px cor_branco">
                    	<div id="txt_1" style="padding-left:20px; float:left; line-height:40px;" align="left"><!--Preencha os campos abaixo--><%=textos_array(43)(2)%></div>
                        <div style="float:right;" align="right"><img id="img_logo_selecionado" src="<%=faixa_logo%>" hspace="10"></div>
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
    	<div style="background:url(<%=faixa_fundo%>); height:45px; width:100%; margin-top:50px;" id="faixa_dir"></div>
    <!-- Faixa Lateral	 -->
    </td>
  </tr>
  <tr>
    <td align="center">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="center" valign="top">&nbsp;</td>
  </tr>
</table>
</div>
<div style="width: 100%; position: absolute; left:0px; float:left;" id="conteudo">
<table width="870" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="547" height="130" colspan="3">&nbsp;</td>
  </tr>
</table>
    <!-- Form Container -->
    <div id="contForm">
    <!-- Form -->
	<form action="default_temp.asp" method="post" id="prcAcessoTicket" name="prcAcessoTicket">
            <!-- Alert error -->
            <div id="aviso_topo" class="fs_12px arial cor_cinza2">
            	<img src="/img/forms/alert-icon.png" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso">
                &nbsp;<span id="txt_topo"><!--Por favor preencher corretamente os itens em destaque !--><%=textos_array(39)(2)%></span>
			</div>
            <!-- End Alert error -->
            <fieldset style="width: 440px;">
            	<legend>Dados para acesso</legend>
				<div id="parcAssis" class="div_parceria" style="height:60px; width:380px;">
                	<label style="width:120px;">LOGIN / CPF
                	  <input type="text" name="frmID_Visitante" id="frmID_Visitante" style="width:110px" max="11" maxlength="11"/></label>
                    <label style="width:120px;">CÓD. IDENTIFICAÇÃO<input type="password" name="frmSenha" id="frmSenha" style="width:110px" max="12" maxlength="12"/></label>
                    <label style="width:20px;">&nbsp;<img class="cursor" src="<%=textos_array(40)(3)%>" onclick="Enviar()" style="padding-top:4px;"/></label>
                </div>
            </fieldset>
            <fieldset style="width: 200px; float:left;">
            	<legend>Lembrar Código de Identificação</legend>
                <div class="div_parceria" style="width:240px; height:60px; background-color: #ccc;">
                	<label style="width:120px;">LOGIN / CPF
                	  <input type="text" name="frmLoginRecuperar" id="frmLoginRecuperar" style="width:110px" max="11" maxlength="11"/></label>
                    <label style="width:20px;">&nbsp;<img class="cursor" src="<%=textos_array(40)(3)%>" onclick="senha()" style="padding-top:4px;"/></label>
                </div>
            </fieldset>           
            <!-- Alert error -->
            <div id="aviso" class="fs_12px arial cor_cinza2" style="display:inline-table; margin-top:15px;">
            	<img src="/img/forms/alert-icon.png" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso">
                &nbsp;<span id="txt"><!--Por favor preencher corretamente os itens em destaque !--><%=textos_array(39)(2)%></span>
			</div>
            <!-- End Alert error -->
        </form>
        <!-- Form End -->
	</div>
    <!-- End Form Container -->
<table width="870" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="547" height="50" colspan="3">&nbsp;</td>
  </tr>
</table>
</div>
<div style="width: 100%; position: absolute;float:left; display:none; z-index:100" id="loading">
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center"><img src="/img/geral/ico_ajax-loader.gif" style="opacity:100"></td>
  </tr>
</table>
</div>
</body>
</html>
<%
Conexao.Close
%>