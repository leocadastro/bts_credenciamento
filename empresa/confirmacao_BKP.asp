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
<script language="javascript" src="/js/funcoes_gerais.js"></script>
<!-- Script desta página -->
<script language="javascript" src="confirmacao.js" charset="utf-8"></script>
<script language="javascript">
var idioma_atual = '<%=Session("cliente_idioma")%>';
</script>
<!-- Script desta página FIM -->
<%

'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================

If Request("id_edicao") = "" OR Request("id_idioma") = "" OR Request("id_tipo") = "" Then
	response.Redirect("/?erro=1")
End If

ID_Edicao 			= Request("id_edicao")
Idioma 				= Request("id_idioma")
Session("cliente_idioma") = IDIOMA
TP_Credenciamento 	= Request("id_tipo")

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

	Pagina_ID = 8
	
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
<% If Request("teste") = "s" Then %>
	<!--#include virtual="/includes/exibir_array.asp"-->
<% End IF
	
	' Select IMG Faixa
	SQL_Img_Faixa 	=	"Select " &_
						"	Img_Faixa " &_
						"From Tipo_Credenciamento " &_
						"Where ID_Tipo_Credenciamento = " & TP_Credenciamento
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

	' Buscar imagens da Feira
	SQL_Edicoes_Configuracao = 	"SELECT " &_
								"	EC.Logo_Email, " &_
								"	EC.Logo_Box, " &_
								"	E.Nome_" & SgIdioma & " as Feira, " &_
								"	EE.Ano as Ano " &_
								"FROM " &_
								"	Edicoes_Configuracao as EC " &_
								"INNER JOIN " &_
								"	Eventos_Edicoes as EE " &_
								"	ON EC.ID_Edicao = EE.ID_Edicao " &_
								"INNER JOIN " &_
								"	Eventos as E" &_
								"	ON EE.ID_Evento = E.ID_Evento " &_
								"WHERE " &_
								"	EC.ID_Edicao = " & ID_Edicao & " " &_
								"	AND EC.Ativo = 1"
	'response.write(SQL_Edicoes_Configuracao)
	Set RS_Edicoes_Configuracao = Server.CreateObject("ADODB.Recordset")
	RS_Edicoes_Configuracao.CursorType = 0
	RS_Edicoes_Configuracao.LockType = 1
	RS_Edicoes_Configuracao.Open SQL_Edicoes_Configuracao, Conexao	

ID_Cadastro		= limpar_texto(Request("frmID_Cadastro"))
For i = Len(ID_Cadastro)+1 To 6
	ID_Cadastro = "0" & ID_Cadastro
Next

ID_Empresa		= limpar_texto(Request("frmID_Empresa"))
Nome 			= limpar_texto(Request("frmNome"))

CPF				= limpar_texto(Request("frmCPF"))
CPFMask			= Mid(CPF,1,3) & "." & Mid(CPF,4,3) & "." & Mid(CPF,7,3) & "-" & Mid(CPF,10,2)

' Select de Cargos
Cargo	 		= limpar_texto(Request("frmCargo"))
SQL_Cargo 		= "SELECT " &_
					"	ID_Cargo as Id, " &_
					"	Cargo_" & SgIdioma & " as Cargo " &_
					"FROM Cargo " &_
					"WHERE " &_
					"	Ativo = 1 " &_
					"	AND ID_Cargo = " & Cargo & " "
Set RS_Cargo = Server.CreateObject("ADODB.Recordset")
RS_Cargo.CursorType = 0
RS_Cargo.LockType = 1
RS_Cargo.Open SQL_Cargo, Conexao	

' Select de Departamentos
Departamento 	= limpar_texto(Request("frmDepartamento"))
SQL_Depto 		= "SELECT " &_
					"	ID_Depto as Id, " &_
					"	Depto_" & SgIdioma & " as Depto " &_
					"FROM Depto " &_
					"WHERE " &_
					"	Ativo = 1 " &_
					"	AND ID_Depto = " & Departamento & "  "
Set RS_Depto = Server.CreateObject("ADODB.Recordset")
RS_Depto.CursorType = 0
RS_Depto.LockType = 1
RS_Depto.Open SQL_Depto, Conexao

CNPJ			= limpar_texto(Request("frmCNPJ"))
CNPJMask 		= Mid(CNPJ,1,2) & "." & Mid(CNPJ,3,3) & "." & Mid(CNPJ,6,3) & "/" & Mid(CNPJ,9,4) & "-" & Mid(CNPJ,13,2)
RazaoSocial		= limpar_texto(Request("frmRazaoSocial"))	
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
                    	<div id="txt_1" style="padding-left:20px; float:left; line-height:40px;" align="left"><!--Conclusao--><%=textos_array(0)(2)%></div>
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
<div style="width: 100%; position: absolute; left:0px; float:left; display:none;" id="conteudo">
<table width="870" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="547" height="130" colspan="3">&nbsp;</td>
  </tr>
</table>
    <!-- Form Container -->
    <div id="contBody">
    	<!--<span class="titulo_confirmacao"><%=textos_array(0)(2)%></span>-->
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
			  <tr>
			    <td width="730" valign="top">
			    	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
			      <tr>
			        <td>&nbsp;</td>
			      </tr>
			      <tr>
			        <td class="arial fs_12px" style="line-height:15px; color: #58595B;"><p>
					<% 
						if ID_Edicao = "5" Then
							texto = textos_array(11)(2)
						Else
							texto = textos_array(1)(2)
						End If
					%>
			            <!--Imprima--><%=texto%><br/><br/>
			            <!--Obrigado--><%=textos_array(2)(2)%>
					<% 
						' Verificando se faz parte da parceria com a Almax
						If ID_Edicao = "1" or ID_Edicao = "10" or ID_Edicao = "2" Then
					%>	
						<br/><br/>
						Garanta as melhores condições em sua viagem e hospedagem <a href="http://www.almax.com.br/" target="_blank" class="link_confirmacao">clique aqui.</a>
					<%
						End If
					%>
			        </p></td>
			      </tr>
			      <tr>
			        <td>&nbsp;</td>
			      </tr>
			      <tr>
			        <td>
			        <table border="0" align="center" class="div_parceria confirmacao" style="padding: 5px;">
			          <tr >
			            <td width="24%">
			               <!--C&oacute;digo de identifica&ccedil;&atilde;o:--><strong><%=textos_array(3)(2)%></strong>
			            </td>
			            <td width="76%" style="font-size:16px;"><strong><%=ID_Cadastro%></strong></td>
			          </tr>
			          <!-- ====================================================== -->
			          <tr>
			            <td><!--CPF--><strong><%=textos_array(4)(2)%></strong></td>
			            <td><%=CPFMask%></td>
			          </tr>
			          <!-- ====================================================== -->
			          <tr>
			            <td><!--Nome--><strong><%=textos_array(5)(2)%></strong></td>
			            <td><%=Nome%></td>
			          </tr>
			          <!-- ====================================================== -->
			          <tr>
			            <td><!--Cargo--><strong><%=textos_array(6)(2)%></strong></td>
			            <td><%=RS_Cargo("Cargo")%></td>
			          </tr>
			          <!-- ====================================================== -->
			          <tr>
			            <td><!--Departamento--><strong><%=textos_array(7)(2)%></strong></td>
			            <td><%=RS_Depto("Depto")%></td>
			          </tr>
			          <!-- ====================================================== -->
                      <% If idioma = "1" Then %>
			          <tr>
			            <td><!--CNPJ--><strong><%=textos_array(8)(2)%></strong></td>
			            <td><%=CNPJMask%></td>
			          </tr>
                      <% End If %>
			          <!-- ====================================================== -->
			          <tr>
			            <td><!--Razão Social--><strong><%=textos_array(9)(2)%></strong></td>
			            <td><%=RazaoSocial%></td>
			          </tr>
			          <!-- ====================================================== -->
			        </table>
			        </td>
			      </tr>
			      <tr>
			        <td>&nbsp;</td>
			      </tr>
			      <tr>
			      	<td>
			      		<table width="100%" border="0" cellspacing="2" cellpadding="2" align="center" style="color: #58595B;">
			      			<tr>
			      				<td align="center" style="background-color: #fff; height: 80px;"><img src='http://cred2012.btsinforma.com.br<%=RS_Edicoes_Configuracao("Logo_Email")%>'/></td>
			      			</tr>
			      		</table>
			      	</td>
			      </tr>		
		          </tr>
			      <tr>
			        <td >&nbsp;</td>
			      </tr>
			      <tr>
			        <td class="arial fs_12px" style="line-height:15px; color: #58595B;">
		        	<%
			        	if ID_Edicao = "5" Then
							texto_rodape = textos_array(12)(2)
						Else
							texto_rodape = textos_array(10)(2)
						End If
					%>
			        	<%=texto_rodape%>
			        </td>
			      </tr>
			      <tr>
			        <td>&nbsp;</td>
			      </tr>
			    	</table>
			    </td>
			  </tr>
			</table>
	</div>
    <!-- End Form Container -->
<table width="870" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="547" height="50" colspan="3">&nbsp;</td>
  </tr>
</table>
</div>
</body>
</html>
<%
RS_Edicoes_Configuracao.Close
Conexao.Close
%>