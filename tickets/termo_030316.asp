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

'response.end
'==================================================
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Application("cnn")
'==================================================

If Session("cliente_edicao") = "" OR Session("cliente_idioma") = "" OR Session("cliente_logado") = "" or Session("cliente_visitante") = "" Then
    response.Redirect("http://www.mbxeventos.net/aol3abf2016/")
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

	Pagina_ID 	= 2
	
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
<% If Limpar_texto(Request("teste")) = "s" Then %>
	<!--#include virtual="/includes/exibir_array.asp"-->
<% End IF
	
	' Select IMG Faixa
	SQL_Img_Faixa 	=	"Select " &_
						"	Img_Faixa " &_
						"From Tipo_Credenciamento " &_
						"Where ID_Tipo_Credenciamento = " & ID_TP_Credenciamento
'	response.write(SQL_Img_Faixa & "<br>")
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
	
	
	
	
	
	'Após o comprovante, caso seja realizado, enviar o email de comprovante
	Session("cliente_enviar_email") = 0
	
%>
<script language="javascript">
var idioma_atual 	= '<%=Session("cliente_idioma")%>';
var select       	= '<%=textos_array(36)(2)%>';
var cor_fundo 	 	= '<%=faixa_cor%>';
var tp_formulario 	= '';

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
                    <td width="189" height="45" background="/img/geral/faixa_fundo_esq.gif"><img id="img_faixa_esq" src="/img/geral/tipos/Faixa_Tickets.gif" width="189" height="45" /></td>
                    <td id="img_fundo_selecionado" height="45" background="<%=faixa_fundo%>" class="atencao_13px cor_branco">
                    	<div id="txt_1" style="padding-left:20px; float:left; line-height:40px;" align="left"><!--Preencha os campos abaixo--></div>
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
    
            <!-- Alert error -->
            <div id="aviso_topo" class="fs_12px arial cor_cinza2">
            	<img src="/img/forms/alert-icon.png" alt="Aviso" width="20" height="20" hspace="2" vspace="4" align="absmiddle" title="Aviso">
                &nbsp;<span id="txt_topo"><!--Por favor preencher corretamente os itens em destaque !--><%=textos_array(39)(2)%></span>
			</div>
            <!-- End Alert error -->
            
            <table width="850" border="0" cellpadding="0" cellspacing="0">
                <tr>
                    <td width="800" height="30" bgcolor="#414042" class="arial fs_13px cor_branco" style="padding-left:15px;"><b>Olá</b> <%=Nome_Visitante%></td>
                    <td width="500" height="30" bgcolor="#414042" align="right"><img src="/img/botoes/voltar.gif" width="47" height="15" hspace="5" class="cursor" onClick="link('status.asp');"></td>
                    <td width="50" height="30" bgcolor="#414042" style=" border-left:#ccc 1px solid;" align="right"><img src="/img/botoes/sair.gif" width="47" height="15" class="cursor" onClick="sair();"></td>
                </tr>
            </table>
            <br/>

            	<%
				SQL_Consulta_Pedidos = 	"Select " &_
										"	P.* " &_
										"From " &_
										"	Pedidos As P " &_
										
										"Where " &_
										"	P.ID_Edicao = '" & Session("cliente_edicao") & "' " &_
										"	And P.ID_Rel_Cadastro = '" & Session("cliente_cadastro") & "' " &_
										"	And P.ID_Visitante = '" & Session("cliente_visitante")  & "' " &_
										"	And P.Status_Pedido = 1"
				'Response.Write(SQL_Consulta_Pedidos)
				
				Set RS_Consulta_Pedidos = Server.CreateObject("ADODB.Recordset")
				RS_Consulta_Pedidos.Open SQL_Consulta_Pedidos, Conexao, 3, 3
				
				If Not RS_Consulta_Pedidos.Eof Then
				
					Tickets 		= True
					Numero_Pedido 	= RS_Consulta_Pedidos("Numero_Pedido")
					'Session("pedido") = RS_Lista_Pedidos("Numero_Pedido")
					ID_Pedido 		= RS_Consulta_Pedidos("ID_Pedido")
					Idioma_Pedido	= RS_Consulta_Pedidos("ID_Idioma")
					Valor_Pedido	= FormatNumber(RS_Consulta_Pedidos("Valor_Pedido"),2)
					
				
				Else
					
					Tickets = False
					
				End If
				%>

				<!--#Include virtual="/tickets/menu_lateral.asp"-->
                
        <form action="/tickets/pedido.asp?acao=novo" method="post" id="termos_condicoes" name="termos_condicoes" >
            <fieldset style="float: right; width: 580px;">
            
	            <input type="hidden" value="0" name="aceito" id="aceito">

            	<legend>Termos & Condições</legend>
				<div id="parcAssis" class="div_parceria" style="width: 580px;">
                
                	<div style="width: 100%; height: 250px; overflow: auto; line-height: 150%">
<strong>TERMOS E CONDIÇÕES</strong><br><br>
Antes de finalizar sua compra conheça as políticas e condições de venda de Ingresso para o evento ABF Franchising Expo 2015:
<br><br>
<strong>1 - DO OBJETO</strong><br>
1.1. A compra estará sujeita à disponibilidade de ingressos e à aprovação da operadora de seu cartão de crédito. Não será permitida a entrada de menores de 16 anos desacompanhados.
<br><br>
<strong>2 - DA AQUISIÇÃO DO INGRESSO</strong><br>
2.1. Confira atentamente os dados de seu pedido antes da confirmação de sua compra. Não será permitido cancelamento ou devolução de ingressos. O seu ingresso é um produto único, ou seja, após sua compra ele não estará mais disponível para venda.
<br><br>
<strong>3 - DA DESISTÊNCIA DA COMPRA E/OU CANCELAMENTO</strong><br>
3.1. Em caso de arrependimento do cliente, o reembolso do valor do ingresso será efetuado, desde que esse direito seja requisitado em até 7 (sete) dias da data da compra, até o limite de 48 (quarenta e oito) horas antes do evento.
<br><br>
3.2. O estorno do valor somente será efetuado mediante a apresentação física de um documento escrito a punho pelo solicitante, para o endereço descrito abaixo*.  O estorno na fatura do cartão seguirá as normas de cada operadora/banco emissor, podendo ser creditado na fatura seguinte ou na subsequente, de acordo com a data de fechamento da fatura. No caso de compras realizadas em cartão de débito, o estorno será realizado como depósito em conta corrente informada pelo cliente, no prazo de até 15 (quinze) dias úteis.
<br><br>
* Endereço para envio do documento escrito, em caso de pedido de estorno: <br>
BTS INFORMA FEIRAS EVENTOS E EDITORA LTDA<br>
Rua Bela Cintra, 967, Conj. 112-A – Cerqueira Cesar – São Paulo - SP<br>
CEP: 01415-003 
<br><br>
3.3. Em caso de cancelamento do evento por parte da Organização, o valor do ingresso será devolvido.
<br><br>
<strong>4 – DA RETIRADA DO INGRESSO</strong><br>
4.1. Para sua segurança, todas as compras, via internet, somente poderão ser entregues ao proprietário do número do CPF utilizado na compra (o nome do titular do pedido deve ser o mesmo nome de cadastro do pedido na internet), mediante a apresentação dos seguintes documentos: comprovante de pagamento da compra e CPF ou Carteira de Motorista de cada pessoa, no qual foi adquirido o ingresso. <br><br>
4.2. O cliente deverá retirar os ingressos adquiridos nos postos de atendimento exclusivos para compra de ingresso online, na entrada do evento, localizado no Expo Center Norte - Pavilhões Azul e Branco - Rua José Bernardo Pinto, 333 - Vila Guilherme - SP, sem taxa de cobrança de entrega.<br><br>
4.3. Para retirada dos ingressos, será requisitada a apresentação de um documento (Carteira de Motorista ou CPF) contendo nome do titular e número do CPF.<br><br>
4.4. Para a retirada de ingressos por terceiros, este deverá apresentar instrumento de procuração do titular do pedido com firma reconhecida em cartório e poderes específicos para a retirada dos ingressos que ficará retida, devendo ainda o outorgado/representante comparecer portando documentos originais de identificação.<br><br>
4.5. Ocorrendo à impossibilidade de entrega dos ingressos, por ausência de pessoa habilitada para o recebimento, não ocorrerá qualquer devolução de valores, sendo que o comprador está ciente das Condições Gerais deste.<br><br>
<strong>5 - DAS CONDIÇÕES GERAIS DE USO</strong><br>
5.1. Terceira idade: Pessoas com mais de 60 anos podem adquirir seus ingressos com desconto de 50% de acordo com o Estatuto do Idoso (Lei nº10.741/2003, cap V, art 23). Um terceiro poderá adquirir os ingressos mediante a apresentação do documento (original ou cópia) do idoso.
<br><br>5.2. A meia-entrada somente será vendida na bilheteria do evento mediante apresentação do documento comprobatório pelo próprio cliente. 
<br><br>5.3 Cada ingresso adquirido no valor de: R$ 60,00 com a compra antecipada pelo site ou R$ 70,00 na bilheteria durante a realização do evento, é válido para os 04 dias de feira.
<br><br>5.4 Visando a proteção dos direitos dos clientes e de terceiros, caso ocorra à tentativa ou a efetiva utilização indevida dos serviços de conveniência de compra on-line de ingressos, a proprietária do evento poderá indicar os dados do cliente às autoridades públicas, aos serviços de proteção ao crédito, dentre outros, para início dos procedimentos legais e administrativos cabíveis.<br>
Excluindo-se os casos de ato ilícito, a proprietária do evento se compromete a não divulgar, ceder, vender ou transferir a terceiros os dados pessoais fornecidos pelo cliente.

<br><br>5.5 – Caso qualquer cláusula do presente contrato seja considerada nula ou sem efeito, não resultará na nulidade total do contrato, permanecendo em vigor nas demais cláusulas, permanecendo os direitos e obrigações ora acordados.

<br><br>5.6 – Fica eleito o foro da Comarca de São Paulo, Estado de São Paulo, para dirimir quaisquer dúvidas oriundas do presente contrato, excluindo-se qualquer outro foro, por mais privilegiado que seja.

                    </div>
                    
					<br><br>

					<a href="#_" onclick="aprovacao(0);"><div id="loader_2" class="bt_recusar_termo" style="float: left; display: block">Recusar</div></a>                    
					<a href="#_" onclick="aprovacao(1);"><div id="loader_2" class="bt_aceitar_termo" style="float: right; display: block">Aprovar</div></a>   
                    
                    <br>
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
    
    <script language="javascript">
    	function aprovacao(termo){
			
			document.getElementById("aceito").value = termo;
			document.getElementById("termos_condicoes").submit();
			
			}
    </script>
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