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
<script language="javascript" src="/js/jquery.maskedinput-1.2.2.min.js"></script>
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

If Session("cliente_edicao") = "" OR Session("cliente_idioma") = "" OR Session("cliente_tipo") = "" Then
	response.Redirect("/?erro=1")
End If

ID_Edicao 				= Session("cliente_edicao")
Idioma 					= Session("cliente_idioma")
ID_TP_Credenciamento 	= 13 'Session("cliente_tipo")
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
'	response.write(SQL_Textos)
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
<% If Request("teste") = "s" Then %>
	<!--#include virtual="/includes/exibir_array.asp"-->
<% End IF
	
	' Select IMG Faixa
	SQL_Img_Faixa 	=	"Select " &_
						"	Img_Faixa " &_
						"From Tipo_Credenciamento " &_
						"Where ID_Tipo_Credenciamento = " & ID_TP_Credenciamento
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
					"	ID_Edicao = " & Session("cliente_edicao")
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
If Request.Form("frmCNPJ") <> "" AND Request.Form("frmSenha") <> "" Then

	CNPJ	= Limpar_Texto(Request.Form("frmCNPJ"))
	CNPJ_MASK = CNPJ
	CNPJ 	= Replace(CNPJ,".","")
	CNPJ 	= Replace(CNPJ,"-","")
	CNPJ 	= Replace(CNPJ,"/","")
	Senha	= Limpar_Texto(Request.Form("frmSenha"))

	' Se vier vazio	
	If Len(Trim(CNPJ)) = 0 OR Len(Trim(Senha)) = 0 Then
		erro = "1"
	Else
		' Verificar se existe Cadastro
		SQL_Cadastro_Universidade =	"Select " &_
									"	RC.ID_Relacionamento_Cadastro " &_
									"	,RC.ID_Edicao " &_
									"	,RC.ID_Empresa " &_
									"	,E.CNPJ " &_
									"	,E.Senha " &_
									"	,E.Razao_Social " &_
									"	,V.Nome_Completo " &_
									"	,V.CPF " &_
									"	,V.ID_Cargo " &_
									"	,V.ID_Depto " &_
									"From Relacionamento_Cadastro as RC " &_
									"Inner Join Tipo_Credenciamento as TC ON TC.ID_Tipo_Credenciamento = RC.ID_Tipo_Credenciamento " &_
									"Inner Join Empresas as E ON E.ID_Empresa = RC.ID_Empresa " &_
									"Inner Join Visitantes as V ON V.ID_Visitante = RC.ID_Visitante " &_
									"Where  " &_
									"	RC.ID_Tipo_Credenciamento = 13					/* Universidade	*/ " &_
									"	AND RC.ID_Edicao = " & ID_Edicao & "	/* Edição	*/ " &_
									"	AND TC.ID_Idioma = " & Idioma & "		/* Idioma	*/ " &_
									"	AND E.CNPJ = '" & CNPJ & "'				/* CNPJ	*/ " &_
									"	AND E.Senha = '" & Senha & "'			/* Senha	*/"
	
'response.write(SQL_Cadastro_Universidade)
	
		Set RS_Cadastro_Universidade = Server.CreateObject("ADODB.Recordset")
		RS_Cadastro_Universidade.Open SQL_Cadastro_Universidade, Conexao
		
		If RS_Cadastro_Universidade.BOF or RS_Cadastro_Universidade.EOF Then
			erro = "2"
		Else
			
			%>
			<form id="alunos" name="alunos" method="POST" action="/alunos/cadastrar.asp">
				<input type="hidden" name="id_edicao" 	value="<%=id_edicao%>">
				<input type="hidden" name="id_idioma" 	value="<%=Idioma%>">
				<input type="hidden" name="id_tipo" 	value="13">
			
				<input type="hidden" name="frmID_Cadastro" 	value="<%=RS_Cadastro_Universidade("ID_Relacionamento_Cadastro")%>">    
				<input type="hidden" name="frmID_Empresa" 	value="<%=RS_Cadastro_Universidade("ID_Empresa")%>">
				<input type="hidden" name="frmNome" 		value="<%=RS_Cadastro_Universidade("Nome_Completo")%>">
				<input type="hidden" name="frmCPF" 			value="<%=RS_Cadastro_Universidade("CPF")%>">
				<input type="hidden" name="frmCargo" 		value="<%=RS_Cadastro_Universidade("ID_Cargo")%>">
				<input type="hidden" name="frmDepartamento" value="<%=RS_Cadastro_Universidade("ID_Depto")%>">
				<input type="hidden" name="frmCNPJ"			value="<%=RS_Cadastro_Universidade("CNPJ")%>">
				<input type="hidden" name="frmRazaoSocial" 	value="<%=RS_Cadastro_Universidade("Razao_Social")%>">
			</form>
			<script language="Javascript">
                $(document).ready(function(){
                    $("#alunos").submit();
                });
            </script>
			<%
			RS_Cadastro_Universidade.Close
		End If
	End If
End If
%>
<script language="javascript">
var idioma_atual = '<%=Session("cliente_idioma")%>';
var select       = '<%=textos_array(36)(2)%>';
var cor_fundo 	 = '<%=faixa_cor%>';
var tp_formulario = '';

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
	}
});
</script>
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
                    <td width="189" height="45" background="/img/geral/faixa_fundo_esq.gif"><img id="img_faixa_esq" src="<%=img_faixa%>" width="189" height="45"></td>
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
<div style="width: 100%; position: absolute; left:0px; float:left; display:;" id="conteudo">
<table width="870" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="547" height="130" colspan="3">&nbsp;</td>
  </tr>
</table>
    <!-- Form Container -->
    <div id="contForm">
    <!-- Form -->
	<form action="/alunos/" method="post" id="prcCadEmpresa" name="prcCadEmpresa" >
            <!-- Alert error -->
            <div id="aviso_topo" class="fs_12px arial cor_cinza2">
            	<img src="/img/forms/alert-icon.png" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso">
                &nbsp;<span id="txt_topo"><!--Por favor preencher corretamente os itens em destaque !--><%=textos_array(39)(2)%></span>
			</div>
            <!-- End Alert error -->
            <fieldset>
            	<legend>Cadastramento de Credenciais de Alunos</legend>
				<div id="parcAssis" class="div_parceria" style="height:130px; width:292px;">
                	<label style="width:280px;">CNPJ<input type="text" name="frmCNPJ" id="frmCNPJ" style="width:270px" max="18" maxlength="18" value="<%=CNPJ_MASK%>"/></label>
                    <label style="width:180px;">SENHA<input type="password" name="frmSenha" id="frmSenha" style="width:170px" max="18" maxlength="18"/></label>
                    <label style="width:20px;">&nbsp;<img class="cursor" src="<%=textos_array(40)(3)%>" onclick="Enviar()" style="padding-top:4px;"/></label>
                </div>
            </fieldset>
            <fieldset>
            	<legend>Esqueci a Senha</legend>
                <div class="div_parceria" style="height:30px; width:292px; line-height:30px; text-align:center;">
                	<a href="javascript:senha();" style="color:#006; text-decoration:none;">Recuperar Senha</a>
                </div>
            </fieldset>
            <br/>
            
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