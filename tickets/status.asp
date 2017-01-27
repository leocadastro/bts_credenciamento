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
<script language="javascript" src="/js/funcoes_gerais.js"></script>
<script language="javascript" src="/js/jquery.alerts.js"></script>
<script language="javascript" src="/js/showpage.js"></script>

<!-- Script desta página -->
<script language="javascript" src="cadastrar.js" charset="utf-8"></script>
<!-- Script desta página FIM -->

<script language="javascript">

function troca_email(id_visitante){

	showPage("troca_email_envia_visitante.asp?id_visitante=" + id_visitante, "td_email_visitante", 0);

}

function exibe_campo(id_visitante){

if(document.getElementById("tr_email_visitante"+id_visitante).style.display == "none"){
	document.getElementById("tr_email_visitante"+id_visitante).style.display = "table-row"
}else{
	document.getElementById("tr_email_visitante"+id_visitante).style.display = "none";
}

}

function Altera_Email(id_visitante){


	if(document.getElementById("email_visitante"+ id_visitante).value==""){

		alert("Preencha com um e-mail");
		document.getElementById("email_visitante"+ id_visitante).focus();
		return;

	}else{

		showPage("altera_email_envia_visitante.asp?id_visitante=" + id_visitante + "&email=" + document.getElementById("email_visitante"+ id_visitante).value, "tr_email_visitante"+id_visitante, 1);

	}

}

function comprovante(pedido,visitante) {
	var erros = 0;

	if (erros == 0) {
		show_loading();
		var timeout = setTimeout(
			function (){
				alert('Tempo de resposta de 15 seg. excedido.\n\nFavor tentar novamente ou reiniciar seu processo.\n\nti@btsmedia.biz');
			}
		, 15000);
		$.getJSON('enviar_comprovante.asp?pedido=' + pedido + '&visitante=' + visitante, function(data, textStatus) {
			$("#loading").fadeOut();
			// verificar erro de retorno
			if (textStatus == 'success') {
				clearTimeout(timeout);
			}
			var msg = '';
			var msg_duvida = '<b>Aten&ccedil;&atilde;o</b>:<br>';
			var objeto = 'bt_adicionar';

			switch (data.retorno) {
				case 'login nao cadastrado':
					msg += 'Login não cadastrado.'
					break;
				case 'email enviado':
					msg = data.nome  + '\nE-mail enviado com sucesso. ';
					//ok();
					break;
			}
			if (msg != '') {
				jAlert(msg, 'Aviso');
			}
		});
	} else {
		$('#aviso').hide().fadeIn().fadeOut().fadeIn();
	}
}
</script>
<%
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================

If Limpar_Texto(Len(Trim(Request.Form("frmID_Cadastro")))) > 0 Then
'If Len(Trim(CStr(Session("IRC")))) > 0 Then

	'Session("cliente_cadastro")	 = Limpar_texto(CStr(Session("IRC")))
    'Session("cliente_empresa")      = Limpar_texto(CStr(Session("ID_Empresa")))
    'Session("cliente_visitante")    = Limpar_texto(CStr(Session("ID_Visitante")))
    'Session("cliente_nome")         = Limpar_texto(CStr(Session("Nome_Completo")))
    'Session("cliente_cpf")          = Limpar_texto(CStr(Session("CPF")))


    Session("cliente_cadastro")     = Limpar_texto(Request.Form("frmID_Cadastro"))
    Session("cliente_empresa")      = Limpar_texto(Request.Form("frmID_Empresa"))
    Session("cliente_visitante")    = Limpar_texto(Request.Form("frmID_Visitante"))
    Session("cliente_nome")         = Limpar_texto(Request.Form("frmNome"))
    Session("cliente_cpf")          = Limpar_texto(Request.Form("frmCPF"))


	'Response.Write(Session("cliente_cadastro") + " - ")
	'Response.Write(Session("cliente_empresa") + " - ")
	'Response.Write(Session("cliente_visitante") + " - ")
	'Response.Write(Session("cliente_nome") + " - ")
	'Response.Write(Session("cliente_cpf") + " - ")

End If


'For Each item In Request.Form
'	Response.Write "" & item & " - Value: " & Request.Form(item) & "<BR />"
'Next


If Session("cliente_edicao") = "" OR Session("cliente_idioma") = "" OR Session("cliente_logado") = "" or Session("cliente_visitante") = "" Then
  response.Redirect("http://www.mbxeventos.net/AOLABF2017/")
End If


ID_Edicao               = Session("cliente_edicao")
Idioma                  = Session("cliente_idioma")
ID_TP_Credenciamento    = Session("cliente_tipo")
TP_Formulario           = Session("cliente_formulario")
IRC                     = Session("cliente_cadastro")
ID_Empresa              = Session("cliente_empresa")
ID_Visitante            = Session("cliente_visitante")
Nome_Visitante          = Session("cliente_nome")
CPF_Visitante           = Session("cliente_cpf")


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
<% If Limpar_texto(Request("teste")) = "s" Then %>
    <!--#include virtual="/includes/exibir_array.asp"-->
<% End IF

    ' Select IMG Faixa
    SQL_Img_Faixa   =   "Select " &_
                        "   Img_Faixa " &_
                        "From Tipo_Credenciamento " &_
                        "Where ID_Tipo_Credenciamento = " & ID_TP_Credenciamento
    'response.write(SQL_Img_Faixa & "<br>")
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

    ' =================================================================================================
    '   ALTERAR DAQUI PRA BAIXO - Santiago - 28/01
    ' =================================================================================================

    Pedido = False
    Ticket = False


    ' Verificar se o Visitante jpa esta cadastrado nesta Edição
    SQL_VerificaEdicaoVisitante =   "SELECT " &_
                                    "   * " &_
                                    "FROM " &_
                                    "    Relacionamento_Cadastro " &_
                                    "Where " &_
                                    "   ID_Visitante = " & ID_Visitante & " " &_
                                    "   AND ID_Edicao = " & ID_Edicao & " "

    'response.write("<hr><b>SQL_VerificaEdicaoVisitante:</b><hr><br/> " & SQL_VerificaEdicaoVisitante & "<hr><br/>")
    Set RS_VerificaEdicaoVisitante = Server.CreateObject("ADODB.Recordset")
    RS_VerificaEdicaoVisitante.CursorType = 0
    RS_VerificaEdicaoVisitante.LockType = 1
    RS_VerificaEdicaoVisitante.Open SQL_VerificaEdicaoVisitante, Conexao

    If not RS_VerificaEdicaoVisitante.BOF or not RS_VerificaEdicaoVisitante.EOF Then
        Pedido = true   ' Pedido pode ser realizado
    Else
        Pedido = False  ' Pedido nao pode ser realizado
    End If

    If Pedido = true Then
        ' Verificar se o Usuário ja comprou o ticket
        SQL_VerificaEdicaoTicket = "SELECT " &_
                                    "   * " &_
                                    "FROM " &_
                                    "    Relacionamento_Cadastro " &_
                                    "Where " &_
                                    "   ID_Visitante = " & ID_Visitante & " " &_
                                    "   AND ID_Edicao = " & ID_Edicao & " "

        'response.write("<hr><b>SQL_VerificaEdicaoTicket:</b><hr>" & SQL_VerificaEdicaoTicket & "<hr><br/>")
        Set RS_VerificaEdicaoTicket = Server.CreateObject("ADODB.Recordset")
        RS_VerificaEdicaoTicket.CursorType = 0
        RS_VerificaEdicaoTicket.LockType = 1
        RS_VerificaEdicaoTicket.Open SQL_VerificaEdicaoTicket, Conexao

        If not RS_VerificaEdicaoTicket.BOF or not RS_VerificaEdicaoTicket.EOF Then
            Ticket = True   ' Pedido pode ser realizado
        Else
            Ticket = False  ' Pedido nao pode ser realizado
        End If

    Else

        ' Se ele n


    End If

    'response.write(Pedido)

    ' Verificando se existe Empresa Vinculado ao Cadstro
    If Len(Trim(ID_Empresa)) <> 0 Then

        SQL_Dados =     "Select " &_
                        "   CNPJ " &_
                        "   ,Razao_Social " &_
                        "   ,Nome_Fantasia " &_
                        "From Empresas " &_
                        "Where " &_
                        "   ID_Empresa = " & ID_Empresa & " " &_
                        "   AND ID_Formulario = 1 "
        Set RS_Dados = Server.CreateObject("ADODB.Recordset")
        RS_Dados.Open SQL_Dados, Conexao, 3, 3

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

    End If




'====Verifica se tem algum pedido em aberto
SQL_Verifica_Pedidos_Existentes = 	"Select " &_
									"	* " &_
									"From Pedidos As P " &_
									"Inner Join Pedidos_Carrinho As C " &_
									"	On C.ID_Pedido = P.ID_Pedido " &_
									"Where " &_
									"	P.ID_Rel_Cadastro = " & Session("cliente_cadastro") & " " &_
									"	And P.ID_Visitante = " & Session("cliente_visitante") & " " &_
									"	And P.Status_Pedido = 1 " &_
									"	And P.ID_Edicao = " & Session("cliente_edicao") & ""
Set RS_Verifica_Pedidos_Existentes = Server.CreateObject("ADODB.Recordset")
RS_Verifica_Pedidos_Existentes.Open SQL_Verifica_Pedidos_Existentes, Conexao, 3, 3

If RS_Verifica_Pedidos_Existentes.Eof Then

	Tenho_Pedido_Aberto = False

Else

	Tenho_Pedido_Aberto = True

End If

'Response.Write("Tenho carrinho aberto? " & Tenho_Pedido_Aberto & "<hr>")

'====Verifica se tenho algum carrinho fechado
SQL_Verifica_Pedidos_Existentes = 	"Select " &_
									"	* " &_
									"From Pedidos As P " &_
									"Inner Join Pedidos_Carrinho As C " &_
									"	On C.ID_Pedido = P.ID_Pedido " &_
									"Where " &_
									"	P.ID_Rel_Cadastro = " & Session("cliente_cadastro") & " " &_
									"	And P.ID_Visitante = " & Session("cliente_visitante") & " " &_
									"	And P.Status_Pedido in (2,3) " &_
									"	And P.ID_Edicao = " & Session("cliente_edicao") & ""
Set RS_Verifica_Pedidos_Existentes = Server.CreateObject("ADODB.Recordset")
RS_Verifica_Pedidos_Existentes.Open SQL_Verifica_Pedidos_Existentes, Conexao, 3, 3

If RS_Verifica_Pedidos_Existentes.Eof Then

	Tenho_Pedido_Fechado = False

Else

	Tenho_Pedido_Fechado = True

End If

'Response.Write("Tenho carrinho fechado? " & Tenho_Pedido_Fechado & "<hr>")


'====Verifica se estou dentro de algum carrinho
SQL_Verifica_Pedidos_Existentes = 	"Select " &_
									"	* " &_
									"From Pedidos As P " &_
									"Inner Join Pedidos_Carrinho As C " &_
									"	On C.ID_Pedido = P.ID_Pedido " &_
									"Where " &_
									"	C.ID_Rel_Cadastro = " & Session("cliente_cadastro") & " " &_
									"	And C.ID_Visitante = " & Session("cliente_visitante") & " " &_
									"	And P.Status_Pedido in (2,3) " &_
									"	And P.ID_Edicao = " & Session("cliente_edicao") & ""
Set RS_Verifica_Pedidos_Existentes = Server.CreateObject("ADODB.Recordset")
RS_Verifica_Pedidos_Existentes.Open SQL_Verifica_Pedidos_Existentes, Conexao, 3, 3

If RS_Verifica_Pedidos_Existentes.Eof Then

	Estou_Em_Algum_Carrinho = False

Else

	Estou_Em_Algum_Carrinho = True

End If

'Response.Write("Estou em algum carrinho? " & Estou_Em_Algum_Carrinho & "<hr>")

'Response.End()


If Tenho_Pedido_Aberto = False And Tenho_Pedido_Fechado = False And Estou_Em_Algum_Carrinho = False Then
	Response.Redirect("novo_pedido.asp")
ElseIf Tenho_Pedido_Aberto = True And Tenho_Pedido_Fechado = False And Estou_Em_Algum_Carrinho = False Then
	Response.Redirect("novo_pedido.asp")
End If



'Response.Write(Request.ServerVariables("http_user_agent"))

If InStr(Request.ServerVariables("http_user_agent"),"MSIE 8.0") Then
	Browser = False
Else
	Browser = True
End If

%>
<script language="javascript">
var idioma_atual    = '<%=Session("cliente_idioma")%>';
var select          = '<%=textos_array(36)(2)%>';
var cor_fundo 	 	= '<%=faixa_cor%>';
var tp_formulario   = '';

function Detalhes_Compra(ID){

	$('#tickets_' + ID).toggle();
	$('#historico_' + ID).toggle();

	}
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
                    <td width="189" height="45" background="/img/geral/faixa_fundo_esqs.gif"><img id="img_faixa_esq" src="/img/geral/tipos/Faixa_Tickets.gif" width="189" height="45"></td>
                    <td id="img_fundo_selecionado" height="45" background="<%=faixa_fundo%>" class="atencao_13px cor_branco">
	                        <div id="txt_1" style="padding-left:20px; float:left; line-height:40px;" align="left"></div>
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
    <td width="547" height="130" colspan="3"></td>
  </tr>
</table><br/>
    <!-- Form Container -->
    <div id="contForm">

            <table width="850" border="0" cellpadding="0" cellspacing="0">
                <tr>
                    <td width="800" height="30" bgcolor="#414042" class="arial fs_13px cor_branco" style="padding-left:15px;"><b>Olá</b> <%=Nome_Visitante%></td>
                    <td width="500" height="30" bgcolor="#414042" align="right"></td>
                    <td width="50" height="30" bgcolor="#414042" style=" border-left:#ccc 1px solid;" align="right"><img src="/img/botoes/sair.gif" width="47" height="15" class="cursor" onClick="sair();"></td>
                </tr>
            </table>
            <br/>

				<%
				SQL_Consulta_Pedidos = 	"Select " &_
										"	*  " &_
										"From Pedidos " &_
										"Where " &_
										"	ID_Edicao = '" & Session("cliente_edicao") & "' " &_
										"	And ID_Rel_Cadastro = '" & Session("cliente_cadastro") & "' " &_
										"	And ID_Visitante = '" & Session("cliente_visitante")  & "' " &_
										"	And Status_Pedido = 1"
				'Response.Write(SQL_Consulta_Pedidos)

				Set RS_Consulta_Pedidos = Server.CreateObject("ADODB.Recordset")
				RS_Consulta_Pedidos.Open SQL_Consulta_Pedidos, Conexao, 3, 3

				If Not RS_Consulta_Pedidos.Eof Then

					Tickets 		= True
					Numero_Pedido 	= RS_Consulta_Pedidos("Numero_Pedido")
					ID_Pedido 		= RS_Consulta_Pedidos("ID_Pedido")
					Idioma_Pedido	= RS_Consulta_Pedidos("ID_Idioma")
					Valor_Pedido	= FormatNumber(RS_Consulta_Pedidos("Valor_Pedido"),2)

				Else

					Tickets = False

				End If
				%>

                <!--#Include virtual="/tickets/menu_lateral.asp"-->

				<form>
                <fieldset style="float: right; width: 580px;">
                <legend>Meus pedidos</legend>
                <div id="parcAssis" class="div_parceria" style="width:580px; float: right;">
                <%
				'SQL_Lista_Pedidos = "Select " &_
				'					"	P.* " &_
				'					"From Pedidos As P " &_
				'					"Where " &_
				'					"	P.ID_Edicao = '" & Session("cliente_edicao") & "'  " &_
				'					"	And P.ID_Rel_Cadastro = '" & IRC & "'  " &_
				'					"	And P.ID_Visitante = '" & ID_Visitante & "'  " &_
				'					"	And P.Status_Pedido <> 1"



				SQL_Lista_Pedidos = "Select " &_
									"	DISTINCT P.*, nome_completo, email_envia " &_
									"From "&_
									"	Pedidos As P " &_
									"INNER JOIN Pedidos_Carrinho As PC " &_
									"	ON PC.ID_Pedido = P.ID_Pedido Inner join visitantes on visitantes.id_visitante = P.id_visitante " &_
									"Where " &_
									"	P.ID_Edicao = '" & Session("cliente_edicao") & "'  " &_
									"	And (P.ID_Rel_Cadastro = '" & IRC & "'  " &_
									"	Or PC.ID_Rel_Cadastro = '" & IRC & "' )" &_
									"	And (P.ID_Visitante = '" & ID_Visitante & "'  " &_
									"	Or PC.ID_Visitante = '" & ID_Visitante & "' )" &_
									"	And P.Status_Pedido <> 1"


				'Response.Write(SQL_Lista_Pedidos)


				Set RS_Lista_Pedidos = Server.CreateObject("ADODB.Recordset")
				RS_Lista_Pedidos.Open SQL_Lista_Pedidos, Conexao, 3, 3


				If Not RS_Lista_Pedidos.Eof Then
				%>
                <div style="padding: 0px 0 10px; font-weight: 100">
                    Lista de compras:
                </div>
                	<table cellpadding="0" cellspacing="0" width="100%">
					<tr>
					<td colspan="3" id="td_email_visitante">Nome: <%=RS_Lista_Pedidos("nome_completo")%><br>Email Comprovante: <%=RS_Lista_Pedidos("email_envia")%></td>
					<td><img src="/img/geral/icones/ico_editar.gif" border="0" onclick="exibe_campo('<%=RS_Lista_Pedidos("ID_Visitante")%>')" title="Alterar E-mail para envio de Confirmação" alt="Alterar E-mail para envio de Confirmação" style="cursor: pointer;"/><input type="hidden" id="id_visitante" value="<%=RS_Lista_Pedidos("id_visitante")%>"></td>

					</tr>

					<tr id="tr_email_visitante<%=RS_Lista_Pedidos("ID_Visitante")%>" style="display:none;" >
													<td colspan="2"><input type="text" id="email_visitante<%=RS_Lista_Pedidos("ID_Visitante")%>" name="email_visitante<%=RS_Lista_Pedidos("ID_Visitante")%>" size="50"></td>
													<td  colspan=> <img src="/img/geral/icones/ok.gif" border="0" onclick="Altera_Email('<%=RS_Lista_Pedidos("ID_Visitante")%>')" title="Alterar E-mail para envio de Confirmação" alt="Alterar E-mail para envio de Confirmação" style="cursor: pointer;"/> </td>


												</tr>
												<tr><td><br></td></tr>
                        <tr bgcolor="414042">
                            <td style="padding: 5px; width: 130px; text-align: left; color: #ffffff;">Nº Pedido</td>
                            <td style="padding: 5px; text-align: center; color: #ffffff;">Total de Ingressos neste Pedido</td>
                            <td style="padding: 5px; width: 80px; text-align: center; color: #ffffff;">Valor</td>
                            <td style="padding: 5px; width: 100px; text-align: center; color: #ffffff;">Ação</td>
                        </tr>
						<%

						P = True
						A = True

							While Not RS_Lista_Pedidos.Eof

							If A = True Then
								Cor_Fundo = "e5e5e5"
							Else
								Cor_Fundo = "efefef"
							End If

							A = Not A

							'Response.Write(ID_Visitante&"<br>"&RS_Lista_Pedidos("ID_Visitante"))

							If CStr(ID_Visitante) <> CStr(RS_Lista_Pedidos("ID_Visitante")) Then

								SQL_Tickets = 	"Select " &_
												"	Count(C.ID_Carrinho) As Tickets " &_
												"From  " &_
												"	Pedidos_Carrinho  As C " &_
												"INNER JOIN Pedidos As P " &_
												"	ON P.ID_Pedido = C.ID_Pedido " &_
												"Where " &_
												"	C.ID_Pedido = '" & RS_Lista_Pedidos("ID_Pedido") & "'" &_
												"	And C.ID_Visitante = '" & ID_Visitante & "' " &_
												"	And C.Cancelado = 0"
							Else
								SQL_Tickets = 	"Select " &_
												"	Count(C.ID_Carrinho) As Tickets " &_
												"From  Pedidos_Carrinho  As C " &_
												"Where " &_
												"	C.ID_Pedido = '" & RS_Lista_Pedidos("ID_Pedido") & "' " &_
												"	And C.Cancelado = 0"
							End if

							'Response.Write(SQL_Tickets)
							Set RS_Tickets = Server.CreateObject("ADODB.Recordset")
							RS_Tickets.Open SQL_Tickets, Conexao, 3, 3

							If Not RS_Tickets.Eof Then
								Tickets = RS_Tickets("Tickets")
							End If

							RS_Tickets.Close

							If P = True Then
								P = False
							Else
								Margin_Top = " margin-top: 15px;"
							End If

							'alterar valor
							Valor_Pedido_Total = Cint(Tickets) * 70

						%>
                        <tr>
                        	<td colspan="4">
                            	<table cellpadding="0" cellspacing="0" width="580" style="border-top: 1px dotted #333; <%=Margin_Top%>">
                                    <tr bgcolor="FFFFFF">
                                        <td width="130" style="padding: 5px; width: 130px; text-align: left"><%=RS_Lista_Pedidos("Numero_Pedido")%></td>
                                        <td width="270" style="padding: 5px; text-align: center;"><%=Tickets%></td>
                                        <td width="80" style="padding: 5px; width: 80px; text-align: center"><%If Cint(RS_Lista_Pedidos("ID_Idioma")) = 1 Then Response.Write("R$ ")%><%=FormatNumber(RS_Lista_Pedidos("Valor_Pedido"),2)%></td>
                                        <td width="100" style="padding: 5px; width: 100px; text-align: center">
                                            <%If Browser = True Then%>
                                            <a href="#detalhes" onclick="Detalhes_Compra(<%=RS_Lista_Pedidos("ID_Pedido")%>)"><img src="/img/geral/icones/lupa.png" width="30" title="Detalhes da Compra" alt="Detalhes da Compra"/></a> &nbsp;
                                            <%End If%>
                                           	<a href="recuperar_confirmacao_pagamento.asp?pedido=<%=RS_Lista_Pedidos("Numero_Pedido")%>" target="_blank"><img src="/img/geral/icones/agt_print-48.png" width="30" title="Imprimir Confirmação de Pagamento" alt="Imprimir Confirmação de Pagamento" style="cursor: pointer;"/></a>
                                        </td>



                                    </tr>

                                    <tr>

                                                <% 'Nome de quem o pedido foi relizado caso o Pedido não seja do visitante
												If CStr(ID_Visitante) <> CStr(RS_Lista_Pedidos("ID_Visitante")) Then
													SQL_Pedido_Nome = 	"Select " &_
																		"	V.Nome_Completo, V.ID_Visitante " &_
																		"From " &_
																		"	Pedidos As P " &_
																		"INNER JOIN Visitantes V " &_
																		"	On V.ID_Visitante = P.ID_Visitante " &_
																		"Where " &_
																		"	ID_Pedido = '" & RS_Lista_Pedidos("ID_Pedido") & "'"

													Set RS_Pedido_Nome = Server.CreateObject("ADODB.Recordset")
                                                	RS_Pedido_Nome.Open SQL_Pedido_Nome, Conexao, 3, 3



                                                	%>
                                                    <td colspan='4' style='padding: 5px 0 5px 5px; background:#ffd51f;'> Este Pedido foi realizado por <%=RS_Pedido_Nome("Nome_Completo")%><BR></td><br>
                                                    <%

												End if
                                                %>
                                    </tr>

                                    <tr id="tickets_<%=RS_Lista_Pedidos("ID_Pedido")%>" <%If Browser = True Then%>style="display: none"<%End If%>>
                                        <td colspan="4">
                                            <table cellpadding="0" cellspacing="0" width="100%" style="border-top: 1px dotted #999">
                                                <tr bgcolor="b6b6b6">
                                                    <td colspan="3" style="padding: 5px; width: 175px; text-align: left; color: #414042">Ingressos adquiridos</td>
                                                </tr>
                                                <tr bgcolor="CCCCCC">
                                                    <td style="padding: 5px; width: 400px; text-align: left">Nome Completo</td>
                                                    <td style="padding: 5px; width: 175px; text-align: center">CPF / Passaporte</td>
                                                    <td style="padding: 5px; width: 100px; text-align: center">Ação</td>
                                                </tr>
                                                <%
												If CStr(ID_Visitante) <> CStr(RS_Lista_Pedidos("ID_Visitante")) Then
													SQL_Tickets = 	"Select " &_
																	"	C.ID_Carrinho,  " &_
																	"	C.ID_Visitante,  " &_
																	"	C.ID_Pedido,  " &_
																	"	C.ID_Rel_Cadastro, " &_
																	"	V.Nome_Completo, " &_
																	"	V.CPF, " &_
																	"	V.Passaporte, V.EMAIL_envia " &_
																	"From  Pedidos_Carrinho  As C " &_
																	"Left Join Visitantes As V On V.ID_Visitante = C.ID_Visitante " &_
																	"Where " &_
																	"	C.ID_Pedido = '" & RS_Lista_Pedidos("ID_Pedido") & "'" &_
																	"	And V.ID_Visitante = '" & ID_Visitante & "'"
												Else
													SQL_Tickets = 	"Select " &_
																	"	C.ID_Carrinho,  " &_
																	"	C.ID_Visitante,  " &_
																	"	C.ID_Pedido,  " &_
																	"	C.ID_Rel_Cadastro, " &_
																	"	V.Nome_Completo, " &_
																	"	V.CPF, " &_
																	"	V.Passaporte, V.EMAIL_envia " &_
																	"From  Pedidos_Carrinho  As C " &_
																	"Left Join Visitantes As V On V.ID_Visitante = C.ID_Visitante " &_
																	"Where " &_
																	"	C.ID_Pedido = '" & RS_Lista_Pedidos("ID_Pedido") & "' " &_
																	"	And C.Cancelado = 0"
												End if

												'Response.Write(SQL_Tickets)
                                                Set RS_Tickets = Server.CreateObject("ADODB.Recordset")
                                                RS_Tickets.Open SQL_Tickets, Conexao, 3, 3

                                                If Not RS_Tickets.Eof Then

                                                Z = True

                                                    While Not RS_Tickets.Eof

                                                    If Z = True Then
                                                        Cor_Fundo = "e5e5e5"
                                                    Else
                                                        Cor_Fundo = "efefef"
                                                    End If

                                                    Z = Not Z
                                                %>

                                                <tr bgcolor="<%=Cor_Fundo%>">
                                                    <td style="padding: 5px; width: 400px; text-align: left" id="td_trocado<%=RS_Tickets("ID_Visitante")%>"><%=RS_Tickets("Nome_Completo")%></td>
                                                    <td style="padding: 5px; width: 175px; text-align: center">
                                                    <%
                                                        If Len(Trim(RS_Tickets("CPF"))) > 0 Then
                                                            Response.Write(RS_Tickets("CPF"))
                                                        Else
                                                            Response.Write(RS_Tickets("Passaporte"))
                                                        End If
                                                    %>
                                                    </td>
                                                    <td style="padding: 5px; width: 100px; text-align: center;">

                                                        <img src="/img/geral/icones/email.png" border="0" onclick="comprovante('<%=RS_Lista_Pedidos("Numero_Pedido")%>','<%=RS_Tickets("ID_Visitante")%>')" title="Enviar Confirmação de Pagamento" alt="Enviar Confirmação de Pagamento" style="cursor: pointer;"/>
														                                                    </td>
                                                </tr>


                                                <%

                                                    RS_Tickets.MoveNext
                                                    Wend
                                                End If
                                                RS_Tickets.Close

                                                %>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr id="historico_<%=RS_Lista_Pedidos("ID_Pedido")%>" style="display: none;">
                                        <td colspan="4">
                                            <table cellpadding="0" cellspacing="0" width="100%" style="border-top: 1px dotted #999">
                                                <tr bgcolor="b6b6b6">
                                                    <td colspan="5" style="padding: 3px; width: 175px; text-align: left; color: #414042">Histórico de Pagamento</td>
                                                </tr>
                                                <tr bgcolor="CCCCCC">
                                                    <td style="padding: 5px; width: 100px; text-align: left">Pedido</td>
                                                    <td style="padding: 5px; width: 150px; text-align: center">Transação</td>
                                                    <td style="padding: 5px; width: 150px; text-align: center">Cód. Autorização</td>
                                                    <td style="padding: 5px; width: 80px; text-align: center">Valor</td>
                                                    <td style="padding: 5px; width: 180px; text-align: center">Data e Hora</td>
                                                </tr>
                                                <%

                                                SQL_Tickets_Pagamento =     "Select " &_
                                                                            "   PH.Numero_Pedido, " &_
                                                                            "   PH.Numero_Transacao, " &_
                                                                            "   PH.Codigo_Autorizacao, " &_
                                                                            "   PH.Data_Pagamento, " &_
                                                                            "   P.Valor_Pedido " &_
                                                                            "From " &_
                                                                            "    Pedidos_Historico As PH " &_
                                                                            "INNER JOIN Pedidos as P " &_
                                                                            "    ON P.Numero_Pedido = PH.Numero_Pedido " &_
                                                                            "Where " &_
                                                                            "   PH.Numero_Pedido = '" & RS_Lista_Pedidos("Numero_Pedido") & "'" &_
                                                                            "   AND Status_Pagamento = '1'"
                                                'response.write(SQL_Tickets_Pagamento)
                                                Set RS_Tickets_Pagamento = Server.CreateObject("ADODB.Recordset")
                                                RS_Tickets_Pagamento.Open SQL_Tickets_Pagamento, Conexao, 3, 3

                                                If Not RS_Tickets_Pagamento.Eof Then
                                                    While Not RS_Tickets_Pagamento.EOF
                                                %>
                                                <tr>
                                                    <td style="padding: 5px; width: 100px; text-align: left;"><%=RS_Tickets_Pagamento("Numero_Pedido")%></td>
                                                    <td style="padding: 5px; width: 150px; text-align: center;"><%=RS_Tickets_Pagamento("Numero_Transacao")%></td>
                                                    <td style="padding: 5px; width: 150px; text-align: center;"><%=RS_Tickets_Pagamento("Codigo_Autorizacao")%></td>
                                                    <td style="padding: 5px; width: 80x; text-align: center;">R$ <%=FormatNumber(Valor_Pedido_Total,2)%></td>
                                                    <td style="padding: 5px; width: 180px; text-align: center;"><%=RS_Tickets_Pagamento("Data_Pagamento")%></td>
                                                </tr>
                                                <%
                                                    RS_Tickets_Pagamento.MoveNext
                                                    Wend
                                                End If
                                                RS_Tickets_Pagamento.Close
                                                %>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
							</td>
                        </tr>
                        <%
							RS_Lista_Pedidos.MoveNext
							Wend
						%>
					</table>
                </div>
						<%
						End If
						RS_Lista_Pedidos.Close
						%>
                </fieldset>
                </form>
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
    <!-- container form end -->
<table width="870" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="547" height="50" colspan="3">&nbsp;</td>
  </tr>
</table>
</div>
<!--#include virtual="/includes/janela_duvida.asp"-->
<form id="confirmacao" name="confirmacao" method="POST" action="/universidades/confirmacao.asp">
    <input type="hidden" name="id_edicao" 		value="<%=Limpar_texto(Request("id_edicao"))%>">
    <input type="hidden" name="id_idioma" 		value="<%=Limpar_texto(Request("id_idioma"))%>">
    <input type="hidden" name="id_tipo" 		value="<%=Limpar_texto(Request("id_tipo"))%>">
    <input type="hidden" name="frmID_Cadastro" 	value="<%=Limpar_texto(Request("frmID_Cadastro"))%>">
    <input type="hidden" name="frmID_Empresa" 	value="<%=Limpar_texto(Request("frmID_Empresa"))%>">
    <input type="hidden" name="frmNome" 		value="<%=Limpar_texto(Request("frmNome"))%>">
    <input type="hidden" name="frmCPF"		 	value="<%=Limpar_texto(Request("frmCPF"))%>">
    <input type="hidden" name="frmCargo" 		value="<%=Limpar_texto(Request("frmCargo"))%>">
    <input type="hidden" name="frmDepartamento" value="<%=Limpar_texto(Request("frmDepartamento"))%>">
    <input type="hidden" name="frmCNPJ" 		value="<%=Limpar_texto(Request("frmCNPJ"))%>">
    <input type="hidden" name="frmRazaoSocial" 	value="<%=Limpar_texto(Request("frmRazaoSocial"))%>">
</form>
</body>
</html>
<%
Conexao.Close
%>
