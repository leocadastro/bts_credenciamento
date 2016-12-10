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
<script language="javascript" src="default.js" charset="utf-8"></script>
<script language="javascript">
var idioma_atual = '<%=Session("cliente_idioma")%>';
</script>
<!-- Script desta página FIM -->
<%
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================

If Session("cliente_edicao") = "" OR Session("cliente_idioma") = "" OR Session("cliente_tipo") = "" Then
    response.Redirect("/?erro=1")
End If

ID_Edicao           = Session("cliente_edicao")
Idioma              = Session("cliente_idioma")
TP_Credenciamento   = Session("cliente_tipo")

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
                    "   ID_Idioma = " & idioma &_
                    "   AND ID_Pagina = " & Pagina_ID &_
                    " Order By ID_Texto "
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
                        "Where ID_Tipo_Credenciamento = " & Session("cliente_tipo")
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
                    "   Nome_" & SgIdioma & " AS Evento " &_
                    "FROM Eventos " &_
                    "WHERE ID_Evento = " & ID_Edicao                
    Set RS_Evento = Server.CreateObject("ADODB.Recordset")
    RS_Evento.CursorType = 0
    RS_Evento.LockType = 1
    RS_Evento.Open SQL_Evento, Conexao
    
    Evento = RS_Evento("Evento")
    Rs_Evento.Close

    ' Select de Feiras
    SQL_Feiras =    "SELECT " &_
                    "   * " &_ 
                    "FROM ProdutoFeira " &_
                    "WHERE " &_ 
                    "    Ativo = 1 "
    Set RS_Feiras = Server.CreateObject("ADODB.Recordset")
    RS_Feiras.CursorType = 0
    RS_Feiras.LockType = 1
    RS_Feiras.Open SQL_Feiras, Conexao
    
    ' Select de Anuncios
    SQL_Anuncio =   "SELECT " &_
                    "   * " &_ 
                    "FROM ProdutoAnuncio " &_ 
                    "WHERE " &_ 
                    "    Ativo = 1 "
    Set RS_Anuncio = Server.CreateObject("ADODB.Recordset")
    RS_Anuncio.CursorType = 0
    RS_Anuncio.LockType = 1
    RS_Anuncio.Open SQL_Anuncio, Conexao

    ' Select de Assinatura
    SQL_Assinatura =    "SELECT " &_
                    "   * " &_ 
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
                        "    Ativo = 1 "
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
                    <td width="189" height="45" background="/img/geral/faixa_fundo_esq.gif"><img id="img_faixa_esq" src="<%=img_faixa%>" width="189" height="45"></td>
                    <td id="img_fundo_selecionado" height="45" background="<%=faixa_fundo%>" class="atencao_13px cor_branco">
                        <div id="txt_1" style="padding-left:20px; float:left; line-height:40px;" align="left">Preencha os campos abaixo</div>
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
            <!-- Alert error -->
            <div id="aviso_topo" class="fs_12px arial cor_cinza2">
                <img src="/img/forms/alert-icon.png" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso">
                &nbsp;Por favor preencher os itens em destaque !
            </div>
            <!-- End Alert error -->
            <fieldset>
                <label style="width:450px;">
                    <div style="width:350px">CNPJ</div>
                    <input type="text" name="frmCNPJ" id="frmCNPJ" style="width:290px" max="18" maxlength="18"/><span class="bt_busca" onclick="getCadastroCNPJ()">Buscar Cadastro</span>
                </label>
      		</fieldset>
            <fieldset>
                <label style="width:390px">Razão Social<input type="text" name="frmRazao" id="frmRazao" style="width:380px" max="150" maxlength="150"/></label>
                <label style="width:390px">Sigla<input type="text" name="frmSigla" id="frmSigla" style="width:380px"/></label>
			</fieldset>
            <fieldset>
                <label style="width:310px">
                    <div style="width:110px">CEP</div>
                    <input type="text" name="frmCEP" id="frmCEP" style="width:100px" max="9" maxlength="9"/><span class="bt_busca" onclick="getEndereco()">Buscar CEP</span>
                </label>
            </fieldset>
            <fieldset>
                <label style="width:440px">Endereço<input type="text" name="frmEndereco" id="frmEndereco" style="width:430px" max="200" maxlength="200"/></label>
                <label style="width:100px">Número<input type="text" name="frmNumero" id="frmNumero" style="width:90px" max="20" maxlength="20"/></label>
                <label style="width:230px">Complemento<input type="text" name="frmComplemento" id="frmComplemento" style="width:220px" max="50" maxlength="50"/></label>
                <label style="width:390px">Bairro<input type="text" name="frmBairro" id="frmBairro" style="width:380px" max="200" maxlength="200"/></label>
                <label style="width:390px">Cidade<input type="text" name="frmCidade" id="frmCidade" style="width:380px" max="200" maxlength="200"/></label>
                <label style="width:260px">UF
                    <select name="frmEstado" id="frmEstado" style="width:250px">
                        <option value="-">-- Selecione --</option>
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
                <label style="width:260px">Pa&iacute;s
                    <select  name="frmPais" id="frmPais" style="width:250px">
                        <option value="-">-- Selecione --</option>
                            <%                          
                            'id_pais = 23 'Brasil
                            If not RS_Pais.BOF or not RS_Pais.EOF Then
                                While not RS_Pais.EOF 
                                    selecionado = ""
                                    If Cstr(id_pais) = Cstr(RS_Pais("ID_Pais")) Then selecionado = " selected "
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
                <label style="width:390px">Site<input type="text" name="frmSite" id="frmSite" style="width:380px"/></label>
			</fieldset>
            <fieldset>
            	<legend>Dados para contato</legend>
      		</fieldset>
            <fieldset>
            	<label style="width:390px">Responsável por comunicação<input type="text" name="frmResp" id="frmResp" style="width:380px"/></label>
           	  	<label style="width:390px">Nome no crachá<input type="text" name="frmNmCracha" id="frmNmCracha" style="width:380px"/></label>
                <label style="width:40px">DDI<input type="text" name="frmDDI" id="frmDDI" style="width:30px"/></label>
                <label style="width:40px">DDD<input type="text" name="frmDDD" id="frmDDD" style="width:30px"/></label>
                <label style="width:100px">Telefone<input name="frmTelefone" id="frmTelefone" style="width:90px"/></label>
                <label style="width:100px">Tipo
                    <select name="frmTipo" id="frmTipo" style="width:90px" onchange="TipoTelefone(this.value);">
                        <option value="-" selected="selected">Selecione</option>
                        <%
                            ' Looping de Tipo de Telefone
                            response.write(ComboTipoTel)
                        %>
                    </select>
                </label>
                <div id="RecebeSmS1">
                    <label style="width:130px; padding-top:32px;">Deseja receber SMS?</label>
                    <div id="radio1">
                        <div class="radiopos"><input type="radio" name="frmSMS" id="frmSMS" value="1"/>&nbsp;sim&nbsp;&nbsp;</div>
                        <div class="radiopos"><input type="radio" name="frmSMS" id="frmSMS" value="0"/>&nbsp;não&nbsp;&nbsp;</div>
                    </div>
                </div>
                <div id="Ramal1">
                    <label style="width:60px">Ramal<input name="frmRamal" id="frmRamal" style="width:50px" max="5" maxlength="5"/></label>
                </div>
                <br/>
                <br/>	
                <label style="width:40px">DDI<input type="text" name="frmDDI2" id="frmDDI2" style="width:30px"/></label>
                <label style="width:40px">DDD<input type="text" name="frmDDD2" id="frmDDD2" style="width:30px"/></label>
                <label style="width:100px">Telefone<input type="text" name="frmTelefone2" id="frmTelefone2" style="width:90px"/></label>
                <label style="width:100px">Tipo
                    <select name="frmTipo2" id="frmTipo2" style="width:90px" onchange="TipoTelefone2(this.value);">
                        <option value="-" selected="selected">Selecione</option>
                        <%
                            ' Looping de Tipo de Telefone
                            response.write(ComboTipoTel)
                        %>
                    </select>
                </label>
                <div id="RecebeSmS2">
                    <label style="width:130px; padding-top:32px;">Deseja receber SMS?</label>
                    <div id="radio2">
                        <div class="radiopos"><input type="radio" name="frmSMS2" id="frmSMS2" value="1"/>&nbsp;sim&nbsp;&nbsp;</div>
                        <div class="radiopos"><input type="radio" name="frmSMS2" id="frmSMS2" value="0"/>&nbsp;não&nbsp;&nbsp;</div>
                    </div>
                </div>
                <div id="Ramal2">
                    <label style="width:60px">Ramal<input name="frmRamal2" id="frmRamal2" style="width:50px" max="5" maxlength="5"/></label>
                </div>
                <br/>
                <label style="width:390px;">E-mail<input type="text" name="frmEmail" id="frmEmail" style="width:380px" value=""/></label>
                <label style="width:390px">Confirme seu E-mail<input type="text" name="frmEmailConf" id="frmEmailConf" style="width:380px"/></label>
            </fieldset>
            <fieldset>
            	<label style="width:790px">Sua universidade se interessaria em parcerias com a BTS Informatica?</label><br/>
                	<input type="checkbox" name="frmInteresse" id="frmInteresse" value="1" onclick="verificar_interesse()"/>&nbsp;<span style="line-height:32px;">Feira</span><br/>
                        <div id="parcFeira" class="div_parceria" style="height:235px;">
                        	<ul>
                        	<%
								' Looping Feiras
								If not RS_Feiras.BOF or not RS_Feiras.EOF Then
									z = 0
                                    While not RS_Feiras.EOF
							%>
                            			<li><input type="checkbox" name="frmInteresseFeira" id="frmInteresseFeira" value="<%=RS_Feiras("ID_Feira")%>"/><span style="display:block;" class="cursor" onclick="$('input[name=frmInteresseFeira]')[<%=z%>].click();" onMouseOver="$(this).addClass('bg_checkbox_radio');" onMouseOut="$(this).removeClass('bg_checkbox_radio');">&nbsp;&nbsp;<%=RS_Feiras("Feira_PTB")%></span></li>
                       		<%
                                        z = z + 1
										RS_Feiras.MoveNext()
									Wend
									RS_Feiras.Close
								End if
							%>
                            </ul>
                        </div>
                   	<input type="checkbox" name="frmInteresse" id="frmInteresse" value="2" onclick="verificar_interesse()"/>&nbsp;<span style="line-height:32px;">Anúncios</span><br/>
                        <div id="parcAnuncio" class="div_parceria" style="height:65px;">
                            <ul>
                        	<%
								' Looping Feiras
								If not RS_Anuncio.BOF or not RS_Anuncio.EOF Then
									While not RS_Anuncio.EOF
							
							%>
                            			<li><input type="checkbox" name="frmInteresseAnuncio" id="frmInteresseAnuncio" value="<%=RS_Anuncio("ID_Anuncio")%>"/>&nbsp;<%=RS_Anuncio("Anuncio_PTB")%></span></li>
                       		<%
										RS_Anuncio.MoveNext()
									Wend
									RS_Anuncio.Close
								End if
							%>
                            </ul>
                        </div>
                    <input type="checkbox" name="frmInteresse" id="frmInteresse" value="3" onclick="verificar_interesse()"/>&nbsp;<span style="line-height:32px;">Assinaturas</span><br/>                    		
                        <div id="parcAssis" class="div_parceria" style="height:65px;">
                            <ul>
                        	<%
								' Looping Feiras
								If not RS_Assinatura.BOF or not RS_Assinatura.EOF Then
									While not RS_Assinatura.EOF
							
							%>
                            			<li><input type="checkbox" name="frmInteresseAssinatura" id="frmInteresseAssinatura" value="<%=RS_Assinatura("ID_Assinatura")%>"/>&nbsp;<%=RS_Assinatura("Assinatura_PTB")%></span></li>
                       		<%
										RS_Assinatura.MoveNext()
									Wend
									RS_Assinatura.Close
								End if
							%>
                            </ul>
                        </div>
                <label style="width:790px">Qual setor sua instituição atua?<br/><input name="frmSetor" id="frmSetor" style="width:780px"/></label>
            </fieldset>
			<div id="acSubmit"><img src="/img/forms/bt_send.png" onclick="Enviar()"/></div>
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