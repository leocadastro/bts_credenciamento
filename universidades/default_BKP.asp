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

ID_Edicao           	= Session("cliente_edicao")
Idioma              	= Session("cliente_idioma")
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
    
    ' Select de Eventos
    SQL_Evento  =   "SELECT " &_
                    "   Nome_" & SgIdioma & " AS Evento, " &_
                    "   Ano " &_
                    "FROM Eventos as E " &_
                    "INNER JOIN" &_
                    "   Eventos_Edicoes as EE " &_
                    "ON EE.ID_Evento = E.ID_Evento " &_
                    "WHERE " &_
                    "   E.Ativo = 1 " &_ 
                    "   AND E.ID_Evento = " & ID_Edicao 

    Set RS_Evento = Server.CreateObject("ADODB.Recordset")
    RS_Evento.CursorType = 0
    RS_Evento.LockType = 1
    RS_Evento.Open SQL_Evento, Conexao
    
    Evento = RS_Evento("Evento") & " " & RS_Evento("Ano")
    Rs_Evento.Close

    ' Select de Feiras
    SQL_Feiras =    "SELECT " &_
                    "   ID_Feira as ID " &_ 
                    "   ,Feira_" & SgIdioma & " as Feira " &_ 
                    "FROM ProdutoFeira " &_
                    "WHERE " &_ 
                    "    Ativo = 1 " &_ 
                    "   Order By Ordem "
    Set RS_Feiras = Server.CreateObject("ADODB.Recordset")
    RS_Feiras.CursorType = 0
    RS_Feiras.LockType = 1
    RS_Feiras.Open SQL_Feiras, Conexao
    
    ' Select de Assinatura
    SQL_Assinatura ="SELECT " &_
                    "   ID_Assinatura as ID " &_ 
                    "   ,Assinatura_" & SgIdioma & " as Assinatura " &_
                    "FROM ProdutoAssinatura " &_ 
                    "WHERE " &_ 
                    "    Ativo = 1 "
    Set RS_Assinatura = Server.CreateObject("ADODB.Recordset")
    RS_Assinatura.CursorType = 0
    RS_Assinatura.LockType = 1
    RS_Assinatura.Open SQL_Assinatura, Conexao

    ' Select de Estados
    SQL_Estado =        "SELECT " &_
                        "   ID_UF, " &_ 
                        "   Sigla, " &_ 
                        "   Estado " &_
                        "FROM UF " &_ 
                        "WHERE " &_ 
						"	Ativo = 1 " &_
						"	AND Sigla <> 'EX' " &_
						"ORDER BY Estado "
    Set RS_Estado = Server.CreateObject("ADODB.Recordset")
    RS_Estado.CursorType = 0
    RS_Estado.LockType = 1
    RS_Estado.Open SQL_Estado, Conexao
    
    ' Select de Paises
    SQL_Pais =          "SELECT " &_
                        "   ID_Pais, " &_
                        "   Pais " &_
                        "FROM Pais " &_
                        "WHERE " &_ 
                        "    Ativo = 1 "
    Set RS_Pais = Server.CreateObject("ADODB.Recordset")
    RS_Pais.CursorType = 0
    RS_Pais.LockType = 1
    RS_Pais.Open SQL_Pais, Conexao

    ' Select de Cargos
    SQL_Cargo =         "SELECT " &_
                        "   ID_Cargo as Id, " &_
                        "   Cargo_" & SgIdioma & " as Cargo " &_
                        "FROM Cargo " &_
                        "WHERE " &_
                        "   Ativo = 1 " &_
                        "ORDER by Cargo " 
    Set RS_Cargo = Server.CreateObject("ADODB.Recordset")
    RS_Cargo.CursorType = 0
    RS_Cargo.LockType = 1
    RS_Cargo.Open SQL_Cargo, Conexao    

    ' Select de Departamentos
    SQL_Depto =         "SELECT " &_
                        "   ID_Depto as Id, " &_
                        "   Depto_" & SgIdioma & " as Depto " &_
                        "FROM Depto " &_
                        "WHERE " &_
                        "   Ativo = 1 " &_
                        "ORDER by Depto " 
    Set RS_Depto = Server.CreateObject("ADODB.Recordset")
    RS_Depto.CursorType = 0
    RS_Depto.LockType = 1
    RS_Depto.Open SQL_Depto, Conexao

    ' Select de Tipo de Telefone
    SQL_TipoTel =       "SELECT " &_
                        "   ID_Tipo_Telefone as Id, " &_
                        "   Tipo_" & SgIdioma & " as Tipo " &_
                        "FROM Tipo_Telefone " &_
                        "ORDER by ID_Tipo_Telefone " 
    Set RS_TipoTel = Server.CreateObject("ADODB.Recordset")
    RS_TipoTel.CursorType = 0
    RS_TipoTel.LockType = 1
    RS_TipoTel.Open SQL_TipoTel, Conexao
    
    ' Looping de Tipo de Telefone
    If not RS_TipoTel.BOF or not RS_TipoTel.EOF Then
        While not RS_TipoTel.EOF
               ComboTipoTel = ComboTipoTel & "<option value='" & RS_TipoTel("Id") & "' sigla='" & RS_TipoTel("Id") & "'>" & caixaAltaBaixa("caixa_altabaixa",RS_TipoTel("Tipo")) & " </option>"
            RS_TipoTel.MoveNext()
        Wend
        RS_TipoTel.Close
    End if

    ' Quantidade de Numeros para telefones
    If SgIdioma = "PTB" Then
        MaxTelefone = "9"
    Else    
        MaxTelefone = "11"
    End If
%>
<script language="javascript">
    var idioma_atual = '<%=Session("cliente_idioma")%>';
    var select       = '<%=textos_array(36)(2)%>';
    var cor_fundo 	 	= '<%=faixa_cor%>';
    var	aviso_msg		= 'Cadastro encontrado, ao continuar você estará atualizando seu dados!';
    var aviso_titulo	= 'Atenção - Atualização de Cadastro';
    var aviso_doc_pj    = '<%=textos_array(53)(2)%>';
    var aviso_doc_pf    = '<%=textos_array(53)(2)%>';
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
	<form action="processar.asp" method="post" id="prcCadUniversidade" name="prcCadUniversidade" >
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
            <%
			' ========================================
			' Não exibir campos para internacional
			If SgIdioma = "PTB" Then
			%>
            <fieldset>
                <label style="width:450px;">
                    <div style="width:350px"><!--CNPJ--><%=textos_array(0)(2)%></div>
                    <input type="text" name="frmCNPJ" id="frmCNPJ" style="width:290px" max="18" maxlength="18"/><span id="bt_busca_cnpj" class="bt_busca" onclick="getCadastroCNPJ()">Buscar Cadastro</span>
                </label>
      		</fieldset>
            <%
			End If
			' ========================================
			%>
            <fieldset>
                <label style="width:390px"><!--Razão Social--><%=textos_array(1)(2)%><input type="text" name="frmRazao" id="frmRazao" style="width:380px" max="150" maxlength="150"/></label>
                <label style="width:390px"><!--Sigla--><%=textos_array(41)(2)%><input type="text" name="frmFantasia" id="frmFantasia" style="width:380px"/></label>
			</fieldset>
            <fieldset>
                <label style="width:390px"><!--Reitor--><%=textos_array(44)(2)%><input type="text" name="frmResp" id="frmResp" style="width:380px"/></label>
            </fieldset>
            <fieldset>
                <label style="width:80px"><!--DDI--><%=textos_array(28)(2)%><input type="text" name="frmDDIEmpresa" id="frmDDIEmpresa" style="width:70px" max="3" maxlength="3"/></label>
                    <label style="width:90px"><!--DDD--><%=textos_array(29)(2)%><input type="text" name="frmDDDEmpresa" id="frmDDDEmpresa" style="width:80px" max="3" maxlength="3"/></label>
                    <label style="width:100px"><!--Telefone--><%=textos_array(30)(2)%><input name="frmTelefoneEmpresa" id="frmTelefoneEmpresa" style="width:90px" max="11" maxlength="11"/></label>
                    <label style="width:100px"><!--Tipo--><%=textos_array(31)(2)%>
                        <select name="frmTipoEmpresa" id="frmTipoEmpresa" style="width:90px" onchange="TipoTelefoneEmpresa(this.value);">
                            <option value="-" selected="selected"><!--Selecione-->-- <%=textos_array(36)(2)%> --</option>
                            <%
                                ' Looping de Tipo de Telefone
                                response.write(ComboTipoTel)
                            %>
                        </select>
                    </label>
                    <div id="RecebeSmSEmpresa" style="display: table;">
                        <label style="width:200px; padding-top:32px;"><!--Deseja receber SMS?--><%=textos_array(32)(2)%></label>
                        <div id="radio3">
                            <div class="radiopos"><input type="radio" name="frmSMSEmpresa" id="frmSMSEmpresa" value="1" checked/><!--sim-->&nbsp;<%=textos_array(34)(2)%>&nbsp;&nbsp;</div>
                            <div class="radiopos"><input type="radio" name="frmSMSEmpresa" id="frmSMSEmpresa" value="0"/><!--nao-->&nbsp;<%=textos_array(35)(2)%>&nbsp;&nbsp;</div>
                        </div>
                    </div>
                        <label style="width:60px" id="RamalEmpresa"><!--Ramal--><%=textos_array(33)(2)%><input name="frmRamalEmpresa" id="frmRamalEmpresa" style="width:50px" max="4" maxlength="4"/></label>
                    <br/>
            </fieldset>
            <br/>
            <fieldset>
                <label style="width:500px">
                    <div style="width:110px"><!--CEP--><%=textos_array(6)(2)%></div>
                    <input type="text" name="frmCEP" id="frmCEP" style="width:100px" max="9" maxlength="9"/>
					<%
					' ========================================
					' Não exibir campos para internacional
					If SgIdioma = "PTB" Then
					%>
                    <span class="bt_busca" onclick="getEndereco()">Buscar CEP</span>
                    <% 
                    End If
					' ========================================
                    %>
                </label>
            </fieldset>
            <fieldset>
                <label style="width:440px"><!--Endereço--><%=textos_array(7)(2)%><input type="text" name="frmEndereco" id="frmEndereco" style="width:430px" max="200" maxlength="200"/></label>
                <label style="width:100px"><!--Número--><%=textos_array(8)(2)%><input type="text" name="frmNumero" id="frmNumero" style="width:90px" max="20" maxlength="20"/></label>
                <label style="width:230px"><!--Complemento--><%=textos_array(9)(2)%><input type="text" name="frmComplemento" id="frmComplemento" style="width:220px" max="50" maxlength="50"/></label>
				<label style="width:390px"><!--Bairro--><%=textos_array(10)(2)%><input type="text" name="frmBairro" id="frmBairro" style="width:380px" max="200" maxlength="200"/></label>
                <label style="width:390px"><!--Cidade--><%=textos_array(11)(2)%><input type="text" name="frmCidade" id="frmCidade" style="width:380px" max="200" maxlength="200"/></label>
                <label style="width:260px"><!--UF--><%=textos_array(12)(2)%>
                    <select name="frmEstado" id="frmEstado" style="width:250px">
                        <option value="-"><!--Selecione-->-- <%=textos_array(36)(2)%> --</option>
                            <%                           
                            If not RS_Estado.BOF or not RS_Estado.EOF Then
                                While not RS_Estado.EOF 
                                    %>
                                        <option value="<%=RS_Estado("ID_UF")%>" sigla="<%=RS_Estado("Sigla")%>"><%=RS_Estado("Estado")%></option>
                                    <%
                                    RS_Estado.MoveNext
                                Wend 
                                RS_Estado.Close
                            End If
                            %>
                    </select>
                </label>
                <label style="width:260px"><!--Pa&iacute;s--><%=textos_array(13)(2)%>
                    <select  name="frmPais" id="frmPais" style="width:250px">
                        <option value="-"><!--Selecione-->-- <%=textos_array(36)(2)%> --</option>
                            <%                          
                            If not RS_Pais.BOF or not RS_Pais.EOF Then
                                While not RS_Pais.EOF 
                                    selecionado = ""
                                    If RS_Pais("Pais") = "Brasil" Then selecionado = " selected "
                                    %>
                                        <option value="<%=Ucase(RS_Pais("ID_Pais"))%>" <%=selecionado%> sigla="<%=Ucase(RS_Pais("Pais"))%>"><%=RS_Pais("Pais")%></option>
                                    <%
                                    RS_Pais.MoveNext
                                Wend 
                                RS_Pais.Close
                            End If
                            %>
                        </select>
                </label>
                <label style="width:390px"><!--Site--><%=textos_array(15)(2)%><input type="text" name="frmSite" id="frmSite" style="width:380px"/></label>
                <br/>
                <label style="width:390px"><!--Senha para login-->Senha para login<input type="password" name="frmSenha" id="frmSenha" style="width:380px" maxlength="20"/></label>
            </fieldset>
            <fieldset>
                <legend><!--Dados para contato--><%=textos_array(17)(2)%></legend>
                <label style="width:450px">
                    <div style="width:400px"><!--CPF--><%=textos_array(18)(2)%></div>
                    <input type="text" name="frmCPF" id="frmCPF" style="width:290px" max="14" maxlength="14"/><span id="bt_busca_cpf" class="bt_busca" onclick="getCadastroCPF()"><!--Buscar CPF--><%=textos_array(19)(2)%></span>
                </label>
            </fieldset>
            <fieldset>
                <label style="width:390px"><!--Nome completo--><%=textos_array(20)(2)%><input type="text" name="frmNome" id="frmNome" style="width:380px" max="100" maxlength="100"/></label>
                <label style="width:390px"><!--Nome no crachá--><%=textos_array(21)(2)%><input type="text" name="frmNmCracha" id="frmNmCracha" style="width:380px" max="27" maxlength="27"/></label>
                <label style="width:120px"><!--Data de Nascimento--><%=textos_array(22)(2)%><input type="text" name="frmDtNasc" id="frmDtNasc" style="width:110px" max="10" maxlength="10"/></label>
                <label style="width:150px"><!--Sexo--><%=textos_array(23)(2)%>
                    <select name="frmSexo" id="frmSexo" style="width:140px;">
                        <option value="-" sigla="-" selected="selected"><!--Selecione-->-- <%=textos_array(36)(2)%> --</option>
                        <option value="0" sigla="0"><!--Masculino--><%=textos_array(24)(2)%></option>
                        <option value="1" sigla="1"><!--Feminino--><%=textos_array(25)(2)%></option>
                    </select>
                </label>
                <br/>   
                <label style="width:390px"><!--Cargo--><%=textos_array(26)(2)%>
                    <select name="frmCargo" id="frmCargo" style="width:390px;" onchange="getSubCargo(this.value);TipoCargoOn(this.value);">
                        <option value="-" sigla="-" selected="selected"><!--Selecione-->-- <%=textos_array(36)(2)%> --</option>
                            <%
                                ' Looping Cargo
                                If not RS_Cargo.BOF or not RS_Cargo.EOF Then
                                    While not RS_Cargo.EOF
                                        %>
                                           <option value="<%=RS_Cargo("Id") %>" sigla="<%=RS_Cargo("Id")%>"><%=caixaAltaBaixa("caixa_altabaixa",RS_Cargo("Cargo")) %></option>
                                        <%
                                        RS_Cargo.MoveNext()
                                    Wend
                                    RS_Cargo.Close
                                End if
                            %>
                    </select>
                    <input type="text" name="frmCargoOutros" id="frmCargoOutros" style="width:230px;display:none"/><span class="bt_fechar" id="FecharCargoOutros" style="display:none" onclick="TipoCargoOff()">x</span>
                </label>
                <label id="SubCargo" style="width:390px"><!--Cargo--><%=textos_array(45)(2)%>
                    <select name="frmSubCargo" id="frmSubCargo" style="width:390px;">
                        <option value="-" sigla="-" selected="selected"><!--Selecione-->-- <%=textos_array(36)(2)%> 
                    </select>
                </label>
                <br/>
                <label style="width:390px"><!--Departamento--><%=textos_array(27)(2)%>
                    <select name="frmDepto" id="frmDepto" style="width:390px;" onchange="TipoDeptoOn(this.value);">
                        <option value="-" sigla="-" selected="selected"><!--Selecione-->-- <%=textos_array(36)(2)%> --</option>
                            <%
                                ' Looping de Departamentos
                                If not RS_Depto.BOF or not RS_Depto.EOF Then
                                    While not RS_Depto.EOF
                                        %>
                                           <option value="<%=RS_Depto("Id") %>" sigla="<%=RS_Depto("Id")%>"><%=caixaAltaBaixa("caixa_altabaixa",RS_Depto("Depto")) %></option>
                                        <%
                                        RS_Depto.MoveNext()
                                    Wend
                                    RS_Depto.Close
                                End if
                            %>
                    </select>
                    <input type="text" name="frmDeptoOutros" id="frmDeptoOutros" style="width:380px;display:none"/><span class="bt_fechar" id="FecharDeptoOutros" style="display:none" onclick="TipoDeptoOff()">x</span>
                </label>
                <br/>
                <label style="width:80px"><!--DDI--><%=textos_array(28)(2)%><input type="text" name="frmDDI" id="frmDDI" style="width:70px" max="3" maxlength="3"/></label>
                <label style="width:90px"><!--DDD--><%=textos_array(29)(2)%><input type="text" name="frmDDD" id="frmDDD" style="width:80px" max="3" maxlength="3"/></label>
                <label style="width:100px"><!--Telefone--><%=textos_array(30)(2)%><input name="frmTelefone" id="frmTelefone" style="width:90px" max="<%=MaxTelefone%>" maxlength="<%=MaxTelefone%>"/></label>
                <label style="width:100px"><!--Tipo--><%=textos_array(31)(2)%>
                    <select name="frmTipo" id="frmTipo" style="width:90px" onchange="TipoTelefone(this.value);">
                        <option value="-" selected="selected"><!--Selecione-->-- <%=textos_array(36)(2)%> --</option>
                        <%
                            ' Looping de Tipo de Telefone
                            response.write(ComboTipoTel)
                        %>
                    </select>
                </label>
                <div id="RecebeSmS1" style="display: table;">
                    <label style="width:200px; padding-top:32px;"><!--Deseja receber SMS?--><%=textos_array(32)(2)%></label>
                    <div id="radio1">
                        <div class="radiopos"><input type="radio" name="frmSMS" id="frmSMS" value="1" checked/><!--sim-->&nbsp;<%=textos_array(34)(2)%>&nbsp;&nbsp;</div>
                        <div class="radiopos"><input type="radio" name="frmSMS" id="frmSMS" value="0"/><!--nao-->&nbsp;<%=textos_array(35)(2)%>&nbsp;&nbsp;</div>
                    </div>
                </div>
                    <label style="width:60px" id="Ramal1"><!--Ramal--><%=textos_array(33)(2)%><input name="frmRamal" id="frmRamal" style="width:50px" max="4" maxlength="4"/></label>
                <br/>
                <label style="width:80px"><!--DDI--><%=textos_array(28)(2)%><input type="text" name="frmDDI2" id="frmDDI2" style="width:70px" max="3" maxlength="3"/></label>
                <label style="width:90px"><!--DDD--><%=textos_array(29)(2)%><input type="text" name="frmDDD2" id="frmDDD2" style="width:80px" max="3" maxlength="3"/></label>
                <label style="width:100px"><!--Telefone--><%=textos_array(30)(2)%><input type="text" name="frmTelefone2" id="frmTelefone2" style="width:90px" max="<%=MaxTelefone%>" maxlength="<%=MaxTelefone%>"/></label>
                <label style="width:100px"><!--Tipo--><%=textos_array(31)(2)%>
                    <select name="frmTipo2" id="frmTipo2" style="width:90px" onchange="TipoTelefone2(this.value);">
                        <option value="-" selected="selected"><!--Selecione-->-- <%=textos_array(36)(2)%> --</option>
                        <%
                            ' Looping de Tipo de Telefone
                            response.write(ComboTipoTel)
                        %>
                    </select>
                </label>
               <div id="RecebeSmS2">
                    <label style="width:200px; padding-top:32px;"><!--Deseja receber SMS?--><%=textos_array(32)(2)%></label>
                    <div id="radio2">
                        <div class="radiopos"><input type="radio" name="frmSMS2" id="frmSMS2" value="1" checked/><!--sim-->&nbsp;<%=textos_array(34)(2)%>&nbsp;&nbsp;</div>
                        <div class="radiopos"><input type="radio" name="frmSMS2" id="frmSMS2" value="0"/><!--nao-->&nbsp;<%=textos_array(35)(2)%>&nbsp;&nbsp;</div>
                    </div>
                </div>
                    <label style="width:60px" id="Ramal2"><!--Ramal--><%=textos_array(33)(2)%><input name="frmRamal2" id="frmRamal2" style="width:50px" max="4" maxlength="4"/></label>
                <br/>
                <label style="width:390px;">E-mail<input type="text" name="frmEmail" id="frmEmail" style="width:380px" max="150" maxlength="150"/></label>
                <label style="width:390px;"><!--Confirme seu--><%=textos_array(37)(2)%> E-mail<input type="text" name="frmEmailConf" id="frmEmailConf" style="width:380px" max="150" maxlength="150"/></label>
                <div id="divNewsletter" style="display:block; width:790px; margin-top:30px;">
                    <div id="radio3" style="width:790px;">
                        <div class="radiopos">
                        <input type="checkbox" name="frmNewsletter" id="frmNewsletter" value="1" checked/><!-- Receber Newsletter-->&nbsp;<%=textos_array(38)(2)%>&nbsp;</div>
                    </div>    
                </div>
                <br/>
            </fieldset>
            <fieldset>
                <legend><!--Perguntas--><%=textos_array(46)(2)%></legend>
            </fieldset>
            <fieldset>
            	<label style="width:790px"><!--Sua instituição se interessaria em parcerias com a BTS Informa?--><%=textos_array(49)(2)%></label><br/>
                <label style="width:790px; margin-top:10px;"><!--Feiras--><%=textos_array(52)(2)%></label><br/>
                    <div id="parcFeira" class="div_parceria" style="height:275px;">
                    	<ul>
                    	<%
							' Looping Feiras
							If not RS_Feiras.BOF or not RS_Feiras.EOF Then
								z = 0
                                While not RS_Feiras.EOF
						%>
                        			<li><input type="checkbox" name="frmInteresseFeira" id="frmInteresseFeira" value="<%=RS_Feiras("ID")%>"/><span class="cursor" onclick="checar(<%=z%>,'<%=RS_Feiras("Feira")%>', 'frmInteresseFeira');$('input[name=frmInteresseFeira]')[<%=z%>].click();">&nbsp;&nbsp;<%=RS_Feiras("Feira")%></span></li>
                   		<%
                                    z = z + 1
									RS_Feiras.MoveNext()
								Wend
								RS_Feiras.Close
							End if
						%>
                        </ul>
                    </div>
                <label style="width:790px; margin-top:10px;"><!--Anúncios--><%=textos_array(50)(2)%></label><br/>
                    <div id="parcAnuncio" class="div_parceria" style="height:85px;">
                        <ul>
                        <%
                            ' Looping Feiras
                            If not RS_Assinatura.BOF or not RS_Assinatura.EOF Then
                                z = 0
                                While not RS_Assinatura.EOF
                        
                        %>
                                    <li><input type="checkbox" name="frmInteresseAnuncio" id="frmInteresseAnuncio" value="<%=RS_Assinatura("ID")%>"/><span class="cursor" onclick="checar(<%=z%>,'<%=RS_Assinatura("Assinatura")%>', 'frmInteresseAnuncio');$('input[name=frmInteresseAnuncio]')[<%=z%>].click();">&nbsp;&nbsp;<%=RS_Assinatura("Assinatura")%></span></li>
                        <%
                                    z = z + 1
                                    RS_Assinatura.MoveNext()
                                Wend
                                RS_Assinatura.MoveFirst
                            End if
                        %>
                        </ul>
                    </div>
                <label style="width:790px; margin-top:10px;"><!--Revistas--><%=textos_array(51)(2)%></label><br/>
                    <div id="parcAssis" class="div_parceria" style="height:85px;">
                        <ul>
                        <%
                            ' Looping Feiras
                            If not RS_Assinatura.BOF or not RS_Assinatura.EOF Then
                                z = 0
                                While not RS_Assinatura.EOF
                        
                        %>
                                    <li><input type="checkbox" name="frmInteresseAssinatura" id="frmInteresseAssinatura" value="<%=RS_Assinatura("ID")%>"/><span class="cursor" onclick="checar(<%=z%>,'<%=RS_Assinatura("Assinatura")%>', 'frmInteresseAssinatura');$('input[name=frmInteresseAssinatura]')[<%=z%>].click();">&nbsp;&nbsp;<%=RS_Assinatura("Assinatura")%></span></li>
                        <%
                                    z = z + 1
                                    RS_Assinatura.MoveNext()
                                Wend
                                RS_Assinatura.Close
                            End if
                        %>
                        </ul>
                    </div>
            </fieldset>
            <br/>
            <!-- Alert error -->
            <div id="aviso" class="fs_12px arial cor_cinza2" style="display:inline-table; margin-top:15px;">
                <img src="/img/forms/alert-icon.png" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso">
                  &nbsp;<!--Por favor preencher corretamente os itens em destaque !--><%=textos_array(39)(2)%>
            </div>
            <br /><br />
            <!-- End Alert error -->
			<div id="acSubmit"style="padding-top:15px;"><img style="cursor: pointer;" src="<%=textos_array(40)(3)%>" onclick="Enviar()"/></div>
        </form>
	</div>
    <!-- container form end -->
<table width="870" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="547" height="50" colspan="3">&nbsp;</td>
  </tr>
</table>
</div>
</body>
</html>
<%
Conexao.Close
%>