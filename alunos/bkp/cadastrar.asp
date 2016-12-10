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
<!-- Script desta página -->
<script language="javascript" src="cadastrar.js" charset="utf-8"></script>
<!-- Script desta página FIM -->
<%
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================

'For Each item In Request.Form
'	Response.Write "" & item & " - Value: " & Request.Form(item) & "<BR />"
'Next

Session("cliente_empresa") = Request.Form("frmID_Empresa")
If Session("cliente_edicao") = "" OR Session("cliente_idioma") = "" OR Session("cliente_logado") = "" OR Session("cliente_empresa") = "" Then
    response.Redirect("/?erro=1")
End If

ID_Edicao           	= Session("cliente_edicao")
Idioma              	= Session("cliente_idioma")
ID_TP_Credenciamento 	= 14 ' Alunos
TP_Formulario 			= 6 '  Alunos

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
    
    SQL_Textos  =   " Select " &_
                    "   ID_Texto, " &_
                    "   ID_Tipo, " &_
                    "   Identificacao, " &_
                    "   Texto, " &_
                    "   URL_Imagem " &_
                    " From Paginas_Textos " &_
                    " Where  " &_
                    "   ID_Idioma = " & idioma & " " &_
                    "   AND ID_Pagina = " & Pagina_ID & " " &_
                    " Order By ORDEM "
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

    
'   For i = Lbound(textos_array) to Ubound(textos_array)
'       response.write("[ i: " & i & " ] [ ident: " & textos_array(i)(1) & " ]  [ txt: " & textos_array(i)(2) & " ]  [ img: " & textos_array(i)(3) & " ]<br>")
'   Next
'===========================================================
%>
<% If Request("teste") = "s" Then %>
    <!--#include virtual="/includes/exibir_array.asp"-->
<% End IF
    
    ' Select IMG Faixa
    SQL_Img_Faixa   =   "Select " &_
                        "   Img_Faixa " &_
                        "From Tipo_Credenciamento " &_
                        "Where ID_Tipo_Credenciamento = " & ID_TP_Credenciamento
    Set RS_Img_Faixa = Server.CreateObject("ADODB.Recordset")
    RS_Img_Faixa.CursorType = 0
    RS_Img_Faixa.LockType = 1
    RS_Img_Faixa.Open SQL_Img_Faixa, Conexao
        img_faixa = RS_Img_Faixa("img_faixa")
    RS_Img_Faixa.Close
    
    ' Faixa TOPO
    SQL_Faixa   =   "Select " &_
                    "   Cor, " &_
                    "   Logo_Negativo, " &_
                    "   Faixa_Fundo " &_
                    "From Edicoes_configuracao " &_
                    "Where  " &_
                    "   ID_Edicao = " & Session("cliente_edicao")
    Set RS_Faixa = Server.CreateObject("ADODB.Recordset")
    RS_Faixa.CursorType = 0
    RS_Faixa.LockType = 1
    RS_Faixa.Open SQL_Faixa, Conexao
        
        faixa_cor   = RS_Faixa("cor")
        faixa_logo  = RS_Faixa("logo_negativo")
        faixa_fundo = RS_Faixa("Faixa_Fundo")
    RS_Faixa.Close
	
    ' Select de Estados
    SQL_TipoCredencial =	"SELECT " &_
							"   ID_Universidade_TipoCredencial, " &_ 
							"   Tipo " &_ 
							"FROM Universidade_TipoCredencial "
    Set RS_TipoCredencial = Server.CreateObject("ADODB.Recordset")
    RS_TipoCredencial.CursorType = 0
    RS_TipoCredencial.LockType = 1
    RS_TipoCredencial.Open SQL_TipoCredencial, Conexao

    SQL_Qtde_Preenchida =   "Select " &_
                            "   Count(ID_Universidade_Credencial) as Total " &_
                            "From Universidade_Credenciais " &_
                            "Where  " &_
                            "   ID_Edicao = " & Session("cliente_edicao") & " " &_
                            "   AND ID_Empresa = " & Session("cliente_empresa") & " "
    Set RS_Qtde_Preenchida = Server.CreateObject("ADODB.Recordset")
    RS_Qtde_Preenchida.Open SQL_Qtde_Preenchida, Conexao
    
    Qtde_Preenchida = 0
    If not RS_Qtde_Preenchida.BOF or not RS_Qtde_Preenchida.EOF Then
        Qtde_Preenchida = RS_Qtde_Preenchida("total")
        RS_Qtde_Preenchida.Close
    End If

    SQL_Dados =     "Select " &_
                    "   CNPJ " &_
                    "   ,Razao_Social " &_
                    "   ,Nome_Fantasia " &_
                    "From Empresas " &_
                    "Where " &_
                    "   ID_Empresa = " & Session("cliente_empresa") & " " &_
                    "   AND ID_Formulario = 5 " ' Cadastrado em Universidade
    Set RS_Dados = Server.CreateObject("ADODB.Recordset")
    RS_Dados.Open SQL_Dados, Conexao
    
    CNPJ    = ""
    Razao   = ""
    Sigla   = ""
    If not RS_Dados.BOF or not RS_Dados.EOF Then
        CNPJ        = RS_Dados("cnpj")
        cnpj_mask   = Mid(cnpj,1,2) & "." & Mid(cnpj,3,3) & "." & Mid(cnpj,6,3) & "/" & Mid(cnpj,9,4) & "-" & Mid(cnpj,13,2)
        Razao       = RS_Dados("Razao_Social")
        Sigla       = RS_Dados("Nome_Fantasia")
        RS_Dados.Close
    End If

%>
<script language="javascript">
var idioma_atual = '<%=Session("cliente_idioma")%>';
var select       = '<%=textos_array(36)(2)%>';
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
    <td align="right">&nbsp;</td></td>
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
    <!-- container form start -->
    <div id="contForm">
    <!-- Form -->
	<form action="processar.asp" method="post" id="prcCadEntidade" name="prcCadEntidade" >
    <input type="hidden" id="id_edicao"     name="id_edicao"        value="<%=Session("cliente_edicao")%>" >
    <input type="hidden" id="id_idioma"     name="id_idioma"        value="<%=Session("cliente_idioma")%>" >
    <input type="hidden" id="id_tipo"       name="id_tipo"          value="<%=ID_TP_Credenciamento%>" >
    <input type="hidden" id="origem_cnpj"   name="origem_cnpj"      value="" >
    <input type="hidden" id="origem_cpf"    name="origem_cpf"       value="" >
    <input type="hidden" id="id_empresa"    name="id_empresa"       value="" >
    <input type="hidden" id="id_visitante"  name="id_visitante"     value="" >
            <!-- Alert error -->
            <div id="aviso_topo" class="fs_12px arial cor_cinza2">
                <img src="/img/forms/alert-icon.png" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso">
                &nbsp;<!--Por favor preencher corretamente os itens em destaque !--><%=textos_array(39)(2)%>
            </div>
            <!-- End Alert error -->
            <fieldset>
                <legend>Universidade</legend>
            </fieldset>
            <fieldset>
                <label style="width:450px;">
                    <div style="width:350px"><!--CNPJ--><%=textos_array(0)(2)%></div>
                    <input type="text" name="frmCNPJ" id="frmCNPJ" style="width:290px" max="18" maxlength="18" value="<%=cnpj_mask%>" disabled="disabled"/>
                </label>
      		</fieldset>
            <fieldset>
                <label style="width:390px"><!--Razão Social--><%=textos_array(1)(2)%><input type="text" name="frmRazao" id="frmRazao" style="width:380px" max="150" maxlength="150" value="<%=razao%>" disabled="disabled"/></label>
                <label style="width:390px"><!--Sigla--><%=textos_array(41)(2)%><input type="text" name="frmSigla" id="frmSigla" style="width:380px" value="<%=sigla%>" disabled="disabled"/></label>
			</fieldset>
            <br/>
            <fieldset>
                <legend>Cadastrar Alunos / Professores</legend>
                <br />
                <table width="790" border="0" cellpadding="0" cellspacing="2">
                  <tr>
                    <td width="248" height="30" class="arial fs_12px cor_cinza1 borda_tabela" align="center">Total de Credenciais à Preencher </td>
                    <td width="80" align="center" bgcolor="#e4e5e6" class="arial fs_12px cor_cinza1 b borda_tabela">
                    <%
                        qtde_total = 99
                        response.write(qtde_total)	
                    %>
                    </td>
                    <td width="248" height="30" class="arial fs_12px cor_cinza1 borda_tabela" align="center">Quantidade Preenchida</td>
                    <% 
                    If Qtde_Preenchida < Cint( qtde_total ) Then
                        ' verde ainda pode preencher mais
                        bg = "#00CC66"
                    Else
                        ' vermelha, não tem mais disponível
                        bg = "#990000"
                    End If
                    %>
                    <td width="80" align="center" bgcolor='<%=bg%>' class="arial fs_12px cor_branco b borda_tabela" id="qtde_preenchida_bg">
                        <span id="qtde_preenchida"><%=Qtde_Preenchida%></span></td>
                    </tr>
                </table>
            </fieldset> 
            <fieldset>
                <label style="width:140px">Tipo de Credencial
                    <select name="frmTipo" id="frmTipo" style="width:140px">
                        <option value="-"><!--Selecione-->-- <%=textos_array(36)(2)%> --</option>
                            <%                           
                            If not RS_TipoCredencial.BOF or not RS_TipoCredencial.EOF Then
                                While not RS_TipoCredencial.EOF 
                                    %>
                                        <option value="<%=RS_TipoCredencial("ID_Universidade_TipoCredencial")%>"><%=RS_TipoCredencial("Tipo")%></option>
                                    <%
                                    RS_TipoCredencial.MoveNext
                                Wend 
                                'RS_TipoCredencial.Close
                                RS_TipoCredencial.MoveFirst
                            End If
                            %>
                    </select>
                </label>
                <label style="width:240px">Nome<input type="text" name="frmNome" id="frmNome" style="width:230px" max="100" maxlength="100"/></label>
                <label style="width:240px">Curso<input type="text" name="frmCurso" id="frmCurso" style="width:230px" max="40" maxlength="40"/></label>
                <label style="width:15px;">Ação<span class="bt_adicionar" id="bt_adicionar"></span></label>
            </fieldset>
            <fieldset style="display:none" id="editar">
                <legend>Editar Alunos / Professores</legend>    
                <label style="width:240px">Tipo de Credencial
                    <input type="hidden" id="frmID" name="frmID" value="">
                    <select name="frmTipo2" id="frmTipo2" style="width:240px">
                        <option value="-"><!--Selecione-->-- <%=textos_array(36)(2)%> --</option>
                            <%                           
                            If not RS_TipoCredencial.BOF or not RS_TipoCredencial.EOF Then
                                While not RS_TipoCredencial.EOF 
                                    %>
                                        <option value="<%=RS_TipoCredencial("ID_Universidade_TipoCredencial")%>"><%=RS_TipoCredencial("Tipo")%></option>
                                    <%
                                    RS_TipoCredencial.MoveNext
                                Wend 
                                RS_TipoCredencial.Close
                            End If
                            %>
                    </select>
                </label>
                <label style="width:240px">Nome<input type="text" name="frmEditarNome" id="frmEditarNome" style="width:230px" max="100" maxlength="100"/></label>
                <label style="width:240px">Curso<input type="text" name="frmEditarCurso" id="frmEditarCurso" style="width:230px" max="40" maxlength="40"/></label>
                <label style="width:20px"></label>
                <table width="750" border="0" cellpadding="0" cellspacing="0" style="margin-bottom:10px;">
                    <tr>
                    <td align="right"><a href="javascript:void('');" id="bt_cancelar" class="bt_cancelar">Cancelar</a></td>
                    <td width="180" align="right"><a href="javascript:void('');" id="bt_atualizar" class="bt_atualizar">Atualizar </a></td>
                    </tr>
                </table>
            </fieldset>
            <fieldset>
                <legend>Lista de Credenciais Cadastradas</legend>
                <br />
                <table width="790" border="0" cellpadding="0" cellspacing="2">
                    <tr>
                        <td height="300" id="iframe_content">
                        <table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0" id="loading_iframe">
                          <tr>
                            <td align="center"><img src="/img/geral/ico_ajax-loader.gif"></td>
                          </tr>
                        </table>
                        </td>
                    </tr>
                </table>
            </fieldset>
            <br/>
            <br/>
            <!-- Alert error -->
            <div id="aviso" class="fs_12px arial cor_cinza2" style="display:inline-table; margin-top:15px; margin-bottom:15px;">
                <img src="/img/forms/alert-icon.png" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso">
                  &nbsp;<!--Por favor preencher corretamente os itens em destaque !--><%=textos_array(39)(2)%>
            </div>
            <!-- End Alert error -->
			<div id="acSubmit" style="margin-top:10px; padding-top:10px;"><img src="<%=textos_array(40)(3)%>" onclick="Enviar()"/></div>
        </form>
	</div>
    <!-- container form end -->
<table width="870" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="547" height="50" colspan="3">&nbsp;</td>
  </tr>
</table>
</div>
<!--#include virtual="/includes/janela_duvida.asp"-->
<form id="confirmacao" name="confirmacao" method="POST" action="/universidades/confirmacao.asp">
    <input type="hidden" name="id_edicao" value="<%=Request("id_edicao")%>">
    <input type="hidden" name="id_idioma" value="<%=Request("id_idioma")%>">
    <input type="hidden" name="id_tipo" value="<%=Request("id_tipo")%>">
    <input type="hidden" name="frmID_Cadastro" value="<%=Request("frmID_Cadastro")%>">
    <input type="hidden" name="frmID_Empresa" value="<%=Request("frmID_Empresa")%>">
    <input type="hidden" name="frmNome" value="<%=Request("frmNome")%>">
    <input type="hidden" name="frmCPF" value="<%=Request("frmCPF")%>">
    <input type="hidden" name="frmCargo" value="<%=Request("frmCargo")%>">
    <input type="hidden" name="frmDepartamento" value="<%=Request("frmDepartamento")%>">
    <input type="hidden" name="frmCNPJ" value="<%=Request("frmCNPJ")%>">
    <input type="hidden" name="frmRazaoSocial" value="<%=Request("frmRazaoSocial")%>">
</form>
</body>
</html>
<%
Conexao.Close
%>